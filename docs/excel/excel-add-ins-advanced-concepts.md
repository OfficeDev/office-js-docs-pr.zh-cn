---
title: Excel JavaScript API 高级概念
description: ''
ms.date: 1/18/2018
ms.openlocfilehash: 89db69e124475c882448a2105837787ce2c84753
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437554"
---
# <a name="excel-javascript-api-advanced-concepts"></a><span data-ttu-id="258bb-102">Excel JavaScript API 高级概念</span><span class="sxs-lookup"><span data-stu-id="258bb-102">Excel JavaScript API advanced concepts</span></span>

<span data-ttu-id="258bb-103">本文根据 [Excel JavaScript API 核心概念](excel-add-ins-core-concepts.md)中的信息，介绍了生成适用于 Excel 2016 的复杂加载项所必需的一些更高级的概念。</span><span class="sxs-lookup"><span data-stu-id="258bb-103">This article builds upon the information in [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016.</span></span> 

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="258bb-104">适用于 Excel 的 Office.js API</span><span class="sxs-lookup"><span data-stu-id="258bb-104">Office.js APIs for Excel</span></span>

<span data-ttu-id="258bb-105">Excel 加载项通过使用适用于 Office 的 JavaScript API 与 Excel 中的对象进行交互，适用于 Office 的 JavaScript API包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="258bb-105">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="258bb-106">**Excel JavaScript API**：[Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。</span><span class="sxs-lookup"><span data-stu-id="258bb-106">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="258bb-107">**公用 API**：公用 API（也称为[共享 API](https://dev.office.com/reference/add-ins/javascript-api-for-office)）随 Office 2013 一起引入，可用于访问诸如 Word、Excel 和 PowerPoint 等多种主机应用程序通用的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="258bb-107">**Common APIs**: Introduced with Office 2013, the common APIs (also referred to as the [Shared API](https://dev.office.com/reference/add-ins/javascript-api-for-office)) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="258bb-p101">尽管很可能会使用 Excel JavaScript API 开发定目标到 Excel 2016 的加载项的大部分功能，但还可以使用共享 API 中的对象。例如：</span><span class="sxs-lookup"><span data-stu-id="258bb-p101">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016, you'll also use objects in the Shared API. For example:</span></span>

- <span data-ttu-id="258bb-110">[Context](https://dev.office.com/reference/add-ins/shared/context)：**Context** 对象表示加载项的运行时环境，并提供对 API 关键对象的访问权限。</span><span class="sxs-lookup"><span data-stu-id="258bb-110">[Context](https://dev.office.com/reference/add-ins/shared/context): The **Context** object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="258bb-111">它由工作簿配置详细信息（如 `contentLanguage` 和 `officeTheme`）组成，并提供有关加载项的运行时环境（如 `host` 和 `platform`）的信息。</span><span class="sxs-lookup"><span data-stu-id="258bb-111">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="258bb-112">此外，它还提供了 `requirements.isSetSupported()` 方法，可用于检查运行加载项的 Excel 应用程序是否支持指定的要求集。</span><span class="sxs-lookup"><span data-stu-id="258bb-112">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span> 

- <span data-ttu-id="258bb-113">[Document](https://dev.office.com/reference/add-ins/shared/document)：**Document** 对象提供 `getFileAsync()` 方法，用于下载运行加载项的 Excel 文件。</span><span class="sxs-lookup"><span data-stu-id="258bb-113">[Document](https://dev.office.com/reference/add-ins/shared/document): The **Document** object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span> 

## <a name="requirement-sets"></a><span data-ttu-id="258bb-114">要求集</span><span class="sxs-lookup"><span data-stu-id="258bb-114">Requirement sets</span></span>

<span data-ttu-id="258bb-115">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="258bb-115">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="258bb-116">Office 加载项可以执行运行时检查或使用清单中指定的要求集确定 Office 主机是否支持加载项所需的 API。</span><span class="sxs-lookup"><span data-stu-id="258bb-116">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span></span> <span data-ttu-id="258bb-117">要确定每个受支持平台上可用的具体要求集，请参阅 [Excel JavaScript API 要求集](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="258bb-117">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="258bb-118">在运行时检查要求集支持</span><span class="sxs-lookup"><span data-stu-id="258bb-118">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="258bb-119">以下代码示例显示如何确定运行加载项的主机应用程序是否支持指定的 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="258bb-119">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="258bb-120">在清单中定义要求集支持</span><span class="sxs-lookup"><span data-stu-id="258bb-120">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="258bb-121">可以在加载项清单中使用[要求元素](https://dev.office.com/reference/add-ins/manifest/requirements)指定加载项要求激活的最小要求集和/或 API 方法。</span><span class="sxs-lookup"><span data-stu-id="258bb-121">You can use the [Requirements element](https://dev.office.com/reference/add-ins/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="258bb-122">如果 Office 主机或平台不支持清单的 **Requirements** 元素中指定的要求集或 API 方法，该加载项不会在该主机或平台中运行，而且不会显示在“我的加载项”**** 中显示的加载项列表中。</span><span class="sxs-lookup"><span data-stu-id="258bb-122">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span> 

<span data-ttu-id="258bb-123">以下代码示例显示加载项清单中的 **Requirements** 元素，该元素指定应在支持 ExcelApi 要求集版本 1.3 或更高版本的所有 Office 主机应用程序中加载该加载项。</span><span class="sxs-lookup"><span data-stu-id="258bb-123">The following code sample shows the **Requirements** element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="258bb-124">若要让加载项适用于 Office 主机的所有平台（如 Excel for Windows、Excel Online 和 Excel for iPad），建议在运行时检查是否有要求支持，而不是在清单中定义要求集支持。</span><span class="sxs-lookup"><span data-stu-id="258bb-124">To make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="258bb-125">Office.js 公用 API 的要求集</span><span class="sxs-lookup"><span data-stu-id="258bb-125">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="258bb-126">有关公用 API 要求集的信息，请参阅 [Office 公用 API 要求集](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="258bb-126">For information about common API requirement sets, see [Office common API requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="258bb-127">加载对象的属性</span><span class="sxs-lookup"><span data-stu-id="258bb-127">Loading the properties of an object</span></span>

<span data-ttu-id="258bb-128">在 Excel JavaScript 对象上调用 `load()` 方法指示 API 在 `sync()` 方法运行时将对象加载到 JavaScript 内存中。</span><span class="sxs-lookup"><span data-stu-id="258bb-128">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span></span> <span data-ttu-id="258bb-129">方法接受字符串（其中包含要加载的以逗号分隔的属性名称）或对象（指定要加载的属性、分页选项等）。`load()`</span><span class="sxs-lookup"><span data-stu-id="258bb-129">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span> 

> [!NOTE]
> <span data-ttu-id="258bb-130">如果对对象（或集合）调用 `load()` 方法，而未指定任何参数，将会加载对象的所有标量属性（或集合中全部对象的所有标量属性）。</span><span class="sxs-lookup"><span data-stu-id="258bb-130">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span></span> <span data-ttu-id="258bb-131">为了减少 Excel 主机应用程序和加载项之间的数据传输量，应避免在没有明确指定要加载的属性的情况下调用 `load()` 方法。</span><span class="sxs-lookup"><span data-stu-id="258bb-131">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

### <a name="method-details"></a><span data-ttu-id="258bb-132">方法详细信息</span><span class="sxs-lookup"><span data-stu-id="258bb-132">Method details</span></span>

#### <a name="loadparam-object"></a><span data-ttu-id="258bb-133">load(param: object)</span><span class="sxs-lookup"><span data-stu-id="258bb-133">load(param: object)</span></span>

<span data-ttu-id="258bb-134">使用参数指定的属性值和对象值填充在 JavaScript 层中创建的代理对象。</span><span class="sxs-lookup"><span data-stu-id="258bb-134">Fills the proxy object created in JavaScript layer with property and object values specified by the parameters.</span></span>

#### <a name="syntax"></a><span data-ttu-id="258bb-135">语法</span><span class="sxs-lookup"><span data-stu-id="258bb-135">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="258bb-136">参数</span><span class="sxs-lookup"><span data-stu-id="258bb-136">Parameters</span></span>

|<span data-ttu-id="258bb-137">**参数**</span><span class="sxs-lookup"><span data-stu-id="258bb-137">**Parameter**</span></span>|<span data-ttu-id="258bb-138">**类型**</span><span class="sxs-lookup"><span data-stu-id="258bb-138">**Type**</span></span>|<span data-ttu-id="258bb-139">**说明**</span><span class="sxs-lookup"><span data-stu-id="258bb-139">**Description**</span></span>|
|:------------|:-------|:----------|
|`param`|<span data-ttu-id="258bb-140">object</span><span class="sxs-lookup"><span data-stu-id="258bb-140">object</span></span>|<span data-ttu-id="258bb-141">可选。</span><span class="sxs-lookup"><span data-stu-id="258bb-141">Optional.</span></span> <span data-ttu-id="258bb-142">接受参数和关系名称作为逗号分隔的字符串或数组。</span><span class="sxs-lookup"><span data-stu-id="258bb-142">Accepts parameter and relationship names as comma-delimited string or an array.</span></span> <span data-ttu-id="258bb-143">也可以传递对象来设置选择和导航属性（如下面的示例所示）。</span><span class="sxs-lookup"><span data-stu-id="258bb-143">An object can also be passed to set the selection and navigation properties (as shown in the example below).</span></span>|

#### <a name="returns"></a><span data-ttu-id="258bb-144">返回</span><span class="sxs-lookup"><span data-stu-id="258bb-144">Returns</span></span>

<span data-ttu-id="258bb-145">void</span><span class="sxs-lookup"><span data-stu-id="258bb-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="258bb-146">示例</span><span class="sxs-lookup"><span data-stu-id="258bb-146">Example</span></span>

<span data-ttu-id="258bb-147">以下代码示例通过复制另一个区域的属性来设置一个 Excel 区域的属性。</span><span class="sxs-lookup"><span data-stu-id="258bb-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span></span> <span data-ttu-id="258bb-148">请注意，必须首先加载源对象，然后才能访问其属性值并将其写入目标区域。</span><span class="sxs-lookup"><span data-stu-id="258bb-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span></span> <span data-ttu-id="258bb-149">此示例假定存在两个区域（**B2:E2** 和 **B7:E7**）的数据，并且这两个区域的初始格式不同。</span><span class="sxs-lookup"><span data-stu-id="258bb-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

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

### <a name="load-option-properties"></a><span data-ttu-id="258bb-150">加载选项属性</span><span class="sxs-lookup"><span data-stu-id="258bb-150">Load option properties</span></span>

<span data-ttu-id="258bb-151">作为调用 `load()` 方法时传递逗号分隔的字符串或数组的替代方法，可以传递一个包含以下属性的对象。</span><span class="sxs-lookup"><span data-stu-id="258bb-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span> 

|<span data-ttu-id="258bb-152">**属性**</span><span class="sxs-lookup"><span data-stu-id="258bb-152">**Property**</span></span>|<span data-ttu-id="258bb-153">**类型**</span><span class="sxs-lookup"><span data-stu-id="258bb-153">**Type**</span></span>|<span data-ttu-id="258bb-154">**说明**</span><span class="sxs-lookup"><span data-stu-id="258bb-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="258bb-155">object</span><span class="sxs-lookup"><span data-stu-id="258bb-155">object</span></span>|<span data-ttu-id="258bb-p109">包含参数/关系名称的逗号分隔列表或数组。可选。</span><span class="sxs-lookup"><span data-stu-id="258bb-p109">Contains a comma-delimited list or an array of parameter/relationship names. Optional.</span></span>|
|`expand`|<span data-ttu-id="258bb-158">object</span><span class="sxs-lookup"><span data-stu-id="258bb-158">object</span></span>|<span data-ttu-id="258bb-p110">包含关系名称的逗号分隔列表或数组。可选。</span><span class="sxs-lookup"><span data-stu-id="258bb-p110">Contains a comma-delimited list or an array of relationship names. Optional.</span></span>|
|`top`|<span data-ttu-id="258bb-161">int</span><span class="sxs-lookup"><span data-stu-id="258bb-161">int</span></span>| <span data-ttu-id="258bb-p111">指定结果中可以包含的集合项最大数量。可选。使用对象表示法选项时，仅可使用此选项。</span><span class="sxs-lookup"><span data-stu-id="258bb-p111">Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="258bb-165">int</span><span class="sxs-lookup"><span data-stu-id="258bb-165">int</span></span>|<span data-ttu-id="258bb-p112">指定要跳过且不包含在结果中的集合中的项数目。如果指定 `top`，跳过指定数目的项目后将会启动结果集。可选。使用对象表示法选项时，仅可使用此选项。</span><span class="sxs-lookup"><span data-stu-id="258bb-p112">Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="258bb-170">以下代码示例通过为集合中的每个工作表的所用区域选择 `name` 属性和 `address` 来加载工作表集合。</span><span class="sxs-lookup"><span data-stu-id="258bb-170">The following code sample loads a workskeet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="258bb-171">它还指定只能加载集合中的前五个工作表。</span><span class="sxs-lookup"><span data-stu-id="258bb-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="258bb-172">可以通过将 `top: 10` 和 `skip: 5` 指定为属性值来处理下一组五个工作表。</span><span class="sxs-lookup"><span data-stu-id="258bb-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span> 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="258bb-173">标量和导航属性</span><span class="sxs-lookup"><span data-stu-id="258bb-173">Scalar and navigation properties</span></span> 

<span data-ttu-id="258bb-174">在 Excel JavaScript API 参考文档中，你可能会注意到，对象成员分为两类：**属性**和**关系**。</span><span class="sxs-lookup"><span data-stu-id="258bb-174">In the Excel JavaScript API reference documentation, you may notice that object members are grouped into two categories: **properties** and **relationships**.</span></span> <span data-ttu-id="258bb-175">对象的属性是一个标量成员（如字符串、整数或布尔值），而对象的关系（也称为“导航属性”）是一个对象/对象集合成员。</span><span class="sxs-lookup"><span data-stu-id="258bb-175">A property of an object is a scalar member such as a string, an integer, or a boolean value, while a relationship of an object (also known as a navigation property) is a member that is either an object or collection of objects.</span></span> <span data-ttu-id="258bb-176">例如，[Worksheet](https://dev.office.com/reference/add-ins/excel/worksheet) 对象中的 `name` 和 `position` 成员是标量属性，而 `protection` 和 `tables` 是关系（导航属性）。</span><span class="sxs-lookup"><span data-stu-id="258bb-176">For example, `name` and `position` members on the [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheet) object are scalar properties, whereas `protection` and `tables` are relationships (navigation properties).</span></span> 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="258bb-177">使用 `object.load()` 的标量属性和导航属性 `object.load()`</span><span class="sxs-lookup"><span data-stu-id="258bb-177">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="258bb-178">调用没有指定参数的 `object.load()` 方法将加载对象的所有标量属性；不会加载对象的导航属性。</span><span class="sxs-lookup"><span data-stu-id="258bb-178">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span></span> <span data-ttu-id="258bb-179">此外，无法直接加载导航属性。</span><span class="sxs-lookup"><span data-stu-id="258bb-179">Additionally, navigation properties cannot be loaded directly.</span></span> <span data-ttu-id="258bb-180">相反，应使用 `load()` 方法引用所需导航属性中的各个标量属性。</span><span class="sxs-lookup"><span data-stu-id="258bb-180">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span></span> <span data-ttu-id="258bb-181">例如，要加载某个区域的字体名称，必须指定 **format** 和 **font** 导航属性作为 **name** 属性的路径：</span><span class="sxs-lookup"><span data-stu-id="258bb-181">For example, to load the font name for a range, you must specify the **format** and **font** navigation properties as the path to the **name** property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="258bb-182">使用 Excel JavaScript API，可以通过遍历路径来设置导航属性的标量属性。</span><span class="sxs-lookup"><span data-stu-id="258bb-182">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="258bb-183">例如，可以使用 `someRange.format.font.size = 10;` 设置区域的字体大小。</span><span class="sxs-lookup"><span data-stu-id="258bb-183">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="258bb-184">在设置之前，不需要加载该属性。</span><span class="sxs-lookup"><span data-stu-id="258bb-184">You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="258bb-185">设置对象的属性</span><span class="sxs-lookup"><span data-stu-id="258bb-185">Setting properties of an object</span></span>

<span data-ttu-id="258bb-186">在具有嵌套导航属性的对象上设置属性可能很麻烦。</span><span class="sxs-lookup"><span data-stu-id="258bb-186">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="258bb-187">作为使用如上所述导航路径设置单个属性的替代方法，可以使用 Excel JavaScript API 中所有对象上可用的 `object.set()` 方法。</span><span class="sxs-lookup"><span data-stu-id="258bb-187">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span></span> <span data-ttu-id="258bb-188">使用此方法，可以通过传递相同 Office.js 类型的另一个对象或 JavaScript 对象（其属性结构类似于调用该方法的对象的属性）一次设置对象的多个属性。</span><span class="sxs-lookup"><span data-stu-id="258bb-188">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="258bb-189">方法只对主机专用 Office JavaScript API（如 Excel JavaScript API）中的对象实现。`set()`</span><span class="sxs-lookup"><span data-stu-id="258bb-189">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span></span> <span data-ttu-id="258bb-190">通用（共享）API 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="258bb-190">The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="258bb-191">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="258bb-191">set (properties: object, options: object)</span></span>

<span data-ttu-id="258bb-p119">对其调用此方法的对象的属性被设置为，由传入对象的相应属性所指定的值。当 `properties` 参数为 JavaScript 对象时，如果传入对象的任何属性与对其调用此方法的对象中的只读属性对应，属性会遭忽略或抛出异常，具体取决于 `options` 参数的值。</span><span class="sxs-lookup"><span data-stu-id="258bb-p119">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="258bb-194">语法</span><span class="sxs-lookup"><span data-stu-id="258bb-194">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="258bb-195">参数</span><span class="sxs-lookup"><span data-stu-id="258bb-195">Parameters</span></span>

|<span data-ttu-id="258bb-196">**参数**</span><span class="sxs-lookup"><span data-stu-id="258bb-196">**Parameter**</span></span>|<span data-ttu-id="258bb-197">**类型**</span><span class="sxs-lookup"><span data-stu-id="258bb-197">**Type**</span></span>|<span data-ttu-id="258bb-198">**说明**</span><span class="sxs-lookup"><span data-stu-id="258bb-198">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="258bb-199">object</span><span class="sxs-lookup"><span data-stu-id="258bb-199">object</span></span>|<span data-ttu-id="258bb-200">与在其上调用方法的对象相同的 Office.js 类型的对象，或属性名称及类型反映在其上调用方法的对象结构的 JavaScript 对象。</span><span class="sxs-lookup"><span data-stu-id="258bb-200">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="258bb-201">object</span><span class="sxs-lookup"><span data-stu-id="258bb-201">object</span></span>|<span data-ttu-id="258bb-p120">可选。只能在首个参数为 JavaScript 对象时传递。此对象可以包含下列属性：`throwOnReadOnly?: boolean`（默认值是 `true`：如果传入的 JavaScript 对象包含只读属性，将引发错误。）</span><span class="sxs-lookup"><span data-stu-id="258bb-p120">Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="258bb-205">返回</span><span class="sxs-lookup"><span data-stu-id="258bb-205">Returns</span></span>

<span data-ttu-id="258bb-206">void</span><span class="sxs-lookup"><span data-stu-id="258bb-206">void</span></span>    

#### <a name="example"></a><span data-ttu-id="258bb-207">示例</span><span class="sxs-lookup"><span data-stu-id="258bb-207">Example</span></span>

<span data-ttu-id="258bb-p121">下面的代码示例设置区域的多个格式属性，具体方法是调用 `set()` 方法，并传入 JavaScript 对象，其中包含可反映 **Range** 对象中属性结构的属性名称和类型。此示例假定区域 **B2:E2** 中有数据。</span><span class="sxs-lookup"><span data-stu-id="258bb-p121">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the **Range** object. This example assumes that there is data in range **B2:E2**.</span></span>

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
## <a name="42ornullobject-methods"></a><span data-ttu-id="258bb-210">\*OrNullObject 方法</span><span class="sxs-lookup"><span data-stu-id="258bb-210">&#42;OrNullObject methods</span></span>

<span data-ttu-id="258bb-211">许多 Excel JavaScript API 方法都会在不符合 API 条件时返回异常。</span><span class="sxs-lookup"><span data-stu-id="258bb-211">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span></span> <span data-ttu-id="258bb-212">例如，如果尝试通过指定工作簿中没有的工作表名称来获取工作表，`getItem()` 方法返回 `ItemNotFound` 异常。</span><span class="sxs-lookup"><span data-stu-id="258bb-212">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="258bb-213">可以使用可用于 Excel JavaScript API 中的多种方法的 `*OrNullObject` 方法变量，而不是为此类应用场景实现复杂的异常处理逻辑。</span><span class="sxs-lookup"><span data-stu-id="258bb-213">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span></span> <span data-ttu-id="258bb-214">方法将返回 null 对象（不是 JavaScript `null`），而不是在指定项不存在的情况下引发异常。`*OrNullObject`</span><span class="sxs-lookup"><span data-stu-id="258bb-214">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span></span> <span data-ttu-id="258bb-215">例如，可以在集合（如 **Worksheets**）上调用 `getItemOrNullObject()` 方法，尝试从集合中检索某个项。</span><span class="sxs-lookup"><span data-stu-id="258bb-215">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span></span> <span data-ttu-id="258bb-216">方法返回指定的项（如果存在）；否则，它将返回 null 对象。`getItemOrNullObject()`</span><span class="sxs-lookup"><span data-stu-id="258bb-216">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span></span> <span data-ttu-id="258bb-217">返回的 null 对象包含布尔属性 `isNullObject`，可以对其进行评估以确定该对象是否存在。</span><span class="sxs-lookup"><span data-stu-id="258bb-217">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="258bb-218">下面的代码示例尝试使用 `getItemOrNullObject()` 方法检索名为“Data”的工作表。</span><span class="sxs-lookup"><span data-stu-id="258bb-218">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="258bb-219">如果此方法返回 null 对象，需要先新建工作表，然后才能对工作表执行操作。</span><span class="sxs-lookup"><span data-stu-id="258bb-219">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="258bb-220">另请参阅</span><span class="sxs-lookup"><span data-stu-id="258bb-220">See also</span></span>
 
* [<span data-ttu-id="258bb-221">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="258bb-221">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="258bb-222">Excel 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="258bb-222">Excel add-ins code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="258bb-223">Excel JavaScript API 性能优化</span><span class="sxs-lookup"><span data-stu-id="258bb-223">Excel JavaScript API performance optimization</span></span>](https://dev.office.com/reference/add-ins/excel/performance.md)
* [<span data-ttu-id="258bb-224">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="258bb-224">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)

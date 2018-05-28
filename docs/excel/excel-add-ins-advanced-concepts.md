---
title: Excel JavaScript API ????
description: ''
ms.date: 1/18/2018
ms.openlocfilehash: 89db69e124475c882448a2105837787ce2c84753
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="excel-javascript-api-advanced-concepts"></a><span data-ttu-id="0feb2-102">Excel JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="0feb2-102">Excel JavaScript API advanced concepts</span></span>

<span data-ttu-id="0feb2-103">???? [Excel JavaScript API ????](excel-add-ins-core-concepts.md)????????????? Excel 2016 ???????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-103">This article builds upon the information in [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016.</span></span> 

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="0feb2-104">??? Excel ? Office.js API</span><span class="sxs-lookup"><span data-stu-id="0feb2-104">Office.js APIs for Excel</span></span>

<span data-ttu-id="0feb2-105">Excel ?????????? Office ? JavaScript API ? Excel ???????????? Office ? JavaScript API???? JavaScript ?????</span><span class="sxs-lookup"><span data-stu-id="0feb2-105">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="0feb2-106">**Excel JavaScript API**?[Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) ? Office 2016 ??????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-106">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="0feb2-107">**?? API**??? API????[?? API](https://dev.office.com/reference/add-ins/javascript-api-for-office)?? Office 2013 ???????????? Word?Excel ? PowerPoint ???????????? UI??????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-107">**Common APIs**: Introduced with Office 2013, the common APIs (also referred to as the [Shared API](https://dev.office.com/reference/add-ins/javascript-api-for-office)) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="0feb2-p101">???????? Excel JavaScript API ?????? Excel 2016 ??????????????????? API ????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-p101">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016, you'll also use objects in the Shared API. For example:</span></span>

- <span data-ttu-id="0feb2-110">[Context](https://dev.office.com/reference/add-ins/shared/context)?**Context** ?????????????????? API ??????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-110">[Context](https://dev.office.com/reference/add-ins/shared/context): The **Context** object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="0feb2-111">????????????? `contentLanguage` ? `officeTheme`???????????????????? `host` ? `platform`?????</span><span class="sxs-lookup"><span data-stu-id="0feb2-111">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="0feb2-112">???????? `requirements.isSetSupported()` ?????????????? Excel ???????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-112">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span> 

- <span data-ttu-id="0feb2-113">[Document](https://dev.office.com/reference/add-ins/shared/document)?**Document** ???? `getFileAsync()` ????????????? Excel ???</span><span class="sxs-lookup"><span data-stu-id="0feb2-113">[Document](https://dev.office.com/reference/add-ins/shared/document): The **Document** object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span> 

## <a name="requirement-sets"></a><span data-ttu-id="0feb2-114">???</span><span class="sxs-lookup"><span data-stu-id="0feb2-114">Requirement sets</span></span>

<span data-ttu-id="0feb2-115">??????????? API ???</span><span class="sxs-lookup"><span data-stu-id="0feb2-115">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="0feb2-116">Office ?????????????????????????? Office ???????????? API?</span><span class="sxs-lookup"><span data-stu-id="0feb2-116">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span></span> <span data-ttu-id="0feb2-117">??????????????????????? [Excel JavaScript API ???](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets)?</span><span class="sxs-lookup"><span data-stu-id="0feb2-117">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="0feb2-118">???????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-118">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="0feb2-119">??????????????????????????????? API ????</span><span class="sxs-lookup"><span data-stu-id="0feb2-119">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="0feb2-120">???????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-120">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="0feb2-121">???????????[????](https://dev.office.com/reference/add-ins/manifest/requirements)????????????????/? API ???</span><span class="sxs-lookup"><span data-stu-id="0feb2-121">You can use the [Requirements element](https://dev.office.com/reference/add-ins/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="0feb2-122">?? Office ??????????? **Requirements** ?????????? API ??????????????????????????????????****???????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-122">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span> 

<span data-ttu-id="0feb2-123">??????????????? **Requirements** ???????????? ExcelApi ????? 1.3 ???????? Office ??????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-123">The following code sample shows the **Requirements** element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="0feb2-124">????????? Office ????????? Excel for Windows?Excel Online ? Excel for iPad?????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-124">To make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="0feb2-125">Office.js ?? API ????</span><span class="sxs-lookup"><span data-stu-id="0feb2-125">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="0feb2-126">???? API ?????????? [Office ?? API ???](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)?</span><span class="sxs-lookup"><span data-stu-id="0feb2-126">For information about common API requirement sets, see [Office common API requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="0feb2-127">???????</span><span class="sxs-lookup"><span data-stu-id="0feb2-127">Loading the properties of an object</span></span>

<span data-ttu-id="0feb2-128">? Excel JavaScript ????? `load()` ???? API ? `sync()` ??????????? JavaScript ????</span><span class="sxs-lookup"><span data-stu-id="0feb2-128">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span></span> <span data-ttu-id="0feb2-129">???????????????????????????????????????????????`load()`</span><span class="sxs-lookup"><span data-stu-id="0feb2-129">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span> 

> [!NOTE]
> <span data-ttu-id="0feb2-130">???????????? `load()` ???????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-130">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span></span> <span data-ttu-id="0feb2-131">???? Excel ????????????????????????????????????????? `load()` ???</span><span class="sxs-lookup"><span data-stu-id="0feb2-131">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

### <a name="method-details"></a><span data-ttu-id="0feb2-132">??????</span><span class="sxs-lookup"><span data-stu-id="0feb2-132">Method details</span></span>

#### <a name="loadparam-object"></a><span data-ttu-id="0feb2-133">load(param: object)</span><span class="sxs-lookup"><span data-stu-id="0feb2-133">load(param: object)</span></span>

<span data-ttu-id="0feb2-134">????????????????? JavaScript ??????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-134">Fills the proxy object created in JavaScript layer with property and object values specified by the parameters.</span></span>

#### <a name="syntax"></a><span data-ttu-id="0feb2-135">??</span><span class="sxs-lookup"><span data-stu-id="0feb2-135">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="0feb2-136">??</span><span class="sxs-lookup"><span data-stu-id="0feb2-136">Parameters</span></span>

|<span data-ttu-id="0feb2-137">**??**</span><span class="sxs-lookup"><span data-stu-id="0feb2-137">**Parameter**</span></span>|<span data-ttu-id="0feb2-138">**??**</span><span class="sxs-lookup"><span data-stu-id="0feb2-138">**Type**</span></span>|<span data-ttu-id="0feb2-139">**??**</span><span class="sxs-lookup"><span data-stu-id="0feb2-139">**Description**</span></span>|
|:------------|:-------|:----------|
|`param`|<span data-ttu-id="0feb2-140">object</span><span class="sxs-lookup"><span data-stu-id="0feb2-140">object</span></span>|<span data-ttu-id="0feb2-141">???</span><span class="sxs-lookup"><span data-stu-id="0feb2-141">Optional.</span></span> <span data-ttu-id="0feb2-142">???????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-142">Accepts parameter and relationship names as comma-delimited string or an array.</span></span> <span data-ttu-id="0feb2-143">????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-143">An object can also be passed to set the selection and navigation properties (as shown in the example below).</span></span>|

#### <a name="returns"></a><span data-ttu-id="0feb2-144">??</span><span class="sxs-lookup"><span data-stu-id="0feb2-144">Returns</span></span>

<span data-ttu-id="0feb2-145">void</span><span class="sxs-lookup"><span data-stu-id="0feb2-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="0feb2-146">??</span><span class="sxs-lookup"><span data-stu-id="0feb2-146">Example</span></span>

<span data-ttu-id="0feb2-147">??????????????????????? Excel ??????</span><span class="sxs-lookup"><span data-stu-id="0feb2-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span></span> <span data-ttu-id="0feb2-148">??????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span></span> <span data-ttu-id="0feb2-149">????????????**B2:E2** ? **B7:E7**????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

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

### <a name="load-option-properties"></a><span data-ttu-id="0feb2-150">??????</span><span class="sxs-lookup"><span data-stu-id="0feb2-150">Load option properties</span></span>

<span data-ttu-id="0feb2-151">???? `load()` ??????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span> 

|<span data-ttu-id="0feb2-152">**??**</span><span class="sxs-lookup"><span data-stu-id="0feb2-152">**Property**</span></span>|<span data-ttu-id="0feb2-153">**??**</span><span class="sxs-lookup"><span data-stu-id="0feb2-153">**Type**</span></span>|<span data-ttu-id="0feb2-154">**??**</span><span class="sxs-lookup"><span data-stu-id="0feb2-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="0feb2-155">object</span><span class="sxs-lookup"><span data-stu-id="0feb2-155">object</span></span>|<span data-ttu-id="0feb2-p109">????/??????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-p109">Contains a comma-delimited list or an array of parameter/relationship names. Optional.</span></span>|
|`expand`|<span data-ttu-id="0feb2-158">object</span><span class="sxs-lookup"><span data-stu-id="0feb2-158">object</span></span>|<span data-ttu-id="0feb2-p110">????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-p110">Contains a comma-delimited list or an array of relationship names. Optional.</span></span>|
|`top`|<span data-ttu-id="0feb2-161">int</span><span class="sxs-lookup"><span data-stu-id="0feb2-161">int</span></span>| <span data-ttu-id="0feb2-p111">????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-p111">Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="0feb2-165">int</span><span class="sxs-lookup"><span data-stu-id="0feb2-165">int</span></span>|<span data-ttu-id="0feb2-p112">?????????????????????????? `top`?????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-p112">Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="0feb2-170">????????????????????????? `name` ??? `address` ?????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-170">The following code sample loads a workskeet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="0feb2-171">???????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="0feb2-172">????? `top: 10` ? `skip: 5` ??????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span> 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="0feb2-173">???????</span><span class="sxs-lookup"><span data-stu-id="0feb2-173">Scalar and navigation properties</span></span> 

<span data-ttu-id="0feb2-174">? Excel JavaScript API ???????????????????????**??**?**??**?</span><span class="sxs-lookup"><span data-stu-id="0feb2-174">In the Excel JavaScript API reference documentation, you may notice that object members are grouped into two categories: **properties** and **relationships**.</span></span> <span data-ttu-id="0feb2-175">????????????????????????????????????????????????/???????</span><span class="sxs-lookup"><span data-stu-id="0feb2-175">A property of an object is a scalar member such as a string, an integer, or a boolean value, while a relationship of an object (also known as a navigation property) is a member that is either an object or collection of objects.</span></span> <span data-ttu-id="0feb2-176">???[Worksheet](https://dev.office.com/reference/add-ins/excel/worksheet) ???? `name` ? `position` ????????? `protection` ? `tables` ??????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-176">For example, `name` and `position` members on the [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheet) object are scalar properties, whereas `protection` and `tables` are relationships (navigation properties).</span></span> 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="0feb2-177">?? `object.load()` ?????????? `object.load()`</span><span class="sxs-lookup"><span data-stu-id="0feb2-177">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="0feb2-178">????????? `object.load()` ???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-178">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span></span> <span data-ttu-id="0feb2-179">??????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-179">Additionally, navigation properties cannot be loaded directly.</span></span> <span data-ttu-id="0feb2-180">?????? `load()` ???????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-180">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span></span> <span data-ttu-id="0feb2-181">???????????????????? **format** ? **font** ?????? **name** ??????</span><span class="sxs-lookup"><span data-stu-id="0feb2-181">For example, to load the font name for a range, you must specify the **format** and **font** navigation properties as the path to the **name** property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="0feb2-182">?? Excel JavaScript API??????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-182">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="0feb2-183">??????? `someRange.format.font.size = 10;` ??????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-183">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="0feb2-184">???????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-184">You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="0feb2-185">???????</span><span class="sxs-lookup"><span data-stu-id="0feb2-185">Setting properties of an object</span></span>

<span data-ttu-id="0feb2-186">???????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-186">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="0feb2-187">???????????????????????????? Excel JavaScript API ????????? `object.set()` ???</span><span class="sxs-lookup"><span data-stu-id="0feb2-187">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span></span> <span data-ttu-id="0feb2-188">?????????????? Office.js ????????? JavaScript ???????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-188">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="0feb2-189">???????? Office JavaScript API?? Excel JavaScript API????????`set()`</span><span class="sxs-lookup"><span data-stu-id="0feb2-189">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span></span> <span data-ttu-id="0feb2-190">??????API ???????</span><span class="sxs-lookup"><span data-stu-id="0feb2-190">The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="0feb2-191">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="0feb2-191">set (properties: object, options: object)</span></span>

<span data-ttu-id="0feb2-p119">??????????????????????????????????? `properties` ??? JavaScript ???????????????????????????????????????????????????? `options` ?????</span><span class="sxs-lookup"><span data-stu-id="0feb2-p119">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="0feb2-194">??</span><span class="sxs-lookup"><span data-stu-id="0feb2-194">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="0feb2-195">??</span><span class="sxs-lookup"><span data-stu-id="0feb2-195">Parameters</span></span>

|<span data-ttu-id="0feb2-196">**??**</span><span class="sxs-lookup"><span data-stu-id="0feb2-196">**Parameter**</span></span>|<span data-ttu-id="0feb2-197">**??**</span><span class="sxs-lookup"><span data-stu-id="0feb2-197">**Type**</span></span>|<span data-ttu-id="0feb2-198">**??**</span><span class="sxs-lookup"><span data-stu-id="0feb2-198">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="0feb2-199">object</span><span class="sxs-lookup"><span data-stu-id="0feb2-199">object</span></span>|<span data-ttu-id="0feb2-200">?????????????? Office.js ????????????????????????????? JavaScript ???</span><span class="sxs-lookup"><span data-stu-id="0feb2-200">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="0feb2-201">object</span><span class="sxs-lookup"><span data-stu-id="0feb2-201">object</span></span>|<span data-ttu-id="0feb2-p120">??????????? JavaScript ??????????????????`throwOnReadOnly?: boolean`????? `true`?????? JavaScript ????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-p120">Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="0feb2-205">??</span><span class="sxs-lookup"><span data-stu-id="0feb2-205">Returns</span></span>

<span data-ttu-id="0feb2-206">void</span><span class="sxs-lookup"><span data-stu-id="0feb2-206">void</span></span>    

#### <a name="example"></a><span data-ttu-id="0feb2-207">??</span><span class="sxs-lookup"><span data-stu-id="0feb2-207">Example</span></span>

<span data-ttu-id="0feb2-p121">?????????????????????????? `set()` ?????? JavaScript ?????????? **Range** ??????????????????????? **B2:E2** ?????</span><span class="sxs-lookup"><span data-stu-id="0feb2-p121">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the **Range** object. This example assumes that there is data in range **B2:E2**.</span></span>

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
## <a name="42ornullobject-methods"></a><span data-ttu-id="0feb2-210">\*OrNullObject ??</span><span class="sxs-lookup"><span data-stu-id="0feb2-210">&#42;OrNullObject methods</span></span>

<span data-ttu-id="0feb2-211">?? Excel JavaScript API ???????? API ????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-211">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span></span> <span data-ttu-id="0feb2-212">??????????????????????????????`getItem()` ???? `ItemNotFound` ???</span><span class="sxs-lookup"><span data-stu-id="0feb2-212">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="0feb2-213">??????? Excel JavaScript API ??????? `*OrNullObject` ???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-213">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span></span> <span data-ttu-id="0feb2-214">????? null ????? JavaScript `null`?????????????????????`*OrNullObject`</span><span class="sxs-lookup"><span data-stu-id="0feb2-214">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span></span> <span data-ttu-id="0feb2-215">?????????? **Worksheets**???? `getItemOrNullObject()` ???????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-215">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span></span> <span data-ttu-id="0feb2-216">?????????????????????? null ???`getItemOrNullObject()`</span><span class="sxs-lookup"><span data-stu-id="0feb2-216">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span></span> <span data-ttu-id="0feb2-217">??? null ???????? `isNullObject`????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-217">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="0feb2-218">??????????? `getItemOrNullObject()` ???????Data??????</span><span class="sxs-lookup"><span data-stu-id="0feb2-218">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="0feb2-219">??????? null ?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-219">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="0feb2-220">????</span><span class="sxs-lookup"><span data-stu-id="0feb2-220">See also</span></span>
 
* [<span data-ttu-id="0feb2-221">Excel JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="0feb2-221">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="0feb2-222">Excel ????????</span><span class="sxs-lookup"><span data-stu-id="0feb2-222">Excel add-ins code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="0feb2-223">Excel JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="0feb2-223">Excel JavaScript API performance optimization</span></span>](https://dev.office.com/reference/add-ins/excel/performance.md)
* [<span data-ttu-id="0feb2-224">Excel JavaScript API ??</span><span class="sxs-lookup"><span data-stu-id="0feb2-224">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)

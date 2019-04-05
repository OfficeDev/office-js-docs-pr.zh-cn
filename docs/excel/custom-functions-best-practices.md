---
ms.date: 01/08/2019
description: 了解在 Excel 中开发自定义函数的最佳实践。
title: 自定义函数最佳实践（预览）
localization_priority: Normal
ms.openlocfilehash: 4efcd0ba5efb0dc7450192694e8f0750de43b8a8
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477542"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="48851-103">自定义函数最佳实践（预览）</span><span class="sxs-lookup"><span data-stu-id="48851-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="48851-104">本文介绍了在 Excel 中开发自定义函数的最佳实践。</span><span class="sxs-lookup"><span data-stu-id="48851-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="troubleshooting"></a><span data-ttu-id="48851-105">故障排除</span><span class="sxs-lookup"><span data-stu-id="48851-105">Troubleshooting</span></span>

1. <span data-ttu-id="48851-106">如果要在 Windows 版 Office 中测试外接程序，则应启用**[运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)**，以解决外接程序的 XML 清单文件及多个安装和运行时条件问题。</span><span class="sxs-lookup"><span data-stu-id="48851-106">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="48851-107">运行时日志记录将 `console.log` 语句写入日志文件，以帮你发现问题。</span><span class="sxs-lookup"><span data-stu-id="48851-107">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

2. <span data-ttu-id="48851-108">如果一个或多个自定义函数与以前注册的外接程序的自定义函数冲突, 则不会加载外接程序。</span><span class="sxs-lookup"><span data-stu-id="48851-108">Your add-in will not load if one or more custom functions conflicts with a previously registered add-in's custom functions.</span></span> <span data-ttu-id="48851-109">在这种情况下, 您可以删除现有加载项, 或者如果在开发加载项时遇到此错误, 则可以在清单中指定不同的命名空间名称。</span><span class="sxs-lookup"><span data-stu-id="48851-109">In this case, you can either remove the existing add-in, or if you encounter this error while developing an add-in, you can specify a different namespace name in your manifest.</span></span>

3. <span data-ttu-id="48851-110">若要向 Excel 自定义函数团队报告有关此故障排除方法的反馈，请发送团队反馈。</span><span class="sxs-lookup"><span data-stu-id="48851-110">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="48851-111">若要执行此操作，请选择“**文件|反馈|发送哭脸**”。</span><span class="sxs-lookup"><span data-stu-id="48851-111">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="48851-112">发送哭脸将提供必要的日志，以帮助我们了解你遇到的问题。</span><span class="sxs-lookup"><span data-stu-id="48851-112">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="48851-113">将函数名称与 JSON 元数据相关联</span><span class="sxs-lookup"><span data-stu-id="48851-113">Associating function names with JSON metadata</span></span>

<span data-ttu-id="48851-114">如[自定义函数概述](custom-functions-overview.md)文章中所述，自定义函数项目必须包含 JSON 元数据文件和脚本（JavaScript 或 TypeScript）文件才能构成完整的函数。</span><span class="sxs-lookup"><span data-stu-id="48851-114">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="48851-115">若要使函数正常工作, 需要将 id 与 JavaScript 实现相关联。</span><span class="sxs-lookup"><span data-stu-id="48851-115">For a function to work properly, you need to associate the id with the JavaScript implementation.</span></span> <span data-ttu-id="48851-116">请确保存在关联, 否则将不会调用该函数。</span><span class="sxs-lookup"><span data-stu-id="48851-116">Make sure there is an association, otherwise the function will not be called.</span></span>

<span data-ttu-id="48851-117">以下代码示例展示了如何执行此关联操作。</span><span class="sxs-lookup"><span data-stu-id="48851-117">The following code sample shows how to do this association.</span></span> <span data-ttu-id="48851-118">该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。</span><span class="sxs-lookup"><span data-stu-id="48851-118">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="48851-119">在 JavaScript 文件中创建自定义函数和在 JSON 元数据文件中指定相应信息时，请记住以下最佳实践。</span><span class="sxs-lookup"><span data-stu-id="48851-119">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="48851-120">在 JSON 元数据文件中，函数的 `name` 和 `id` 只能使用大写字母。</span><span class="sxs-lookup"><span data-stu-id="48851-120">Only use uppercase letters for a function's `name` and `id` in the JSON metadata file.</span></span> <span data-ttu-id="48851-121">不要使用大小写字母混合或仅使用小写字母。</span><span class="sxs-lookup"><span data-stu-id="48851-121">Do not use a mix of cases or only lowercase letters.</span></span> <span data-ttu-id="48851-122">如果这样做，你最终可能会得到两个值，这些值只会因情况而异，从而导致意外覆盖你的函数。</span><span class="sxs-lookup"><span data-stu-id="48851-122">If you do, you may end up with two values that only differ by case which will cause unintentional overwriting of your functions.</span></span> <span data-ttu-id="48851-123">例如，`id` 值为 **add** 的函数对象稍后可以通过声明在 `id` 值为 **ADD** 的函数对象文件中覆盖。</span><span class="sxs-lookup"><span data-stu-id="48851-123">For example, a function object with an `id` value of **add** could be overwritten by declaration later in the file of function object with an `id` value of **ADD**.</span></span> <span data-ttu-id="48851-124">此外，`name` 属性还会定义最终用户将在 Excel 中看到的函数名称。</span><span class="sxs-lookup"><span data-stu-id="48851-124">Additionally, the `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="48851-125">使用大写字母作为每个自定义函数的名称可在 Excel 中提供一致的体验，其中所有内置函数名称均为大写。</span><span class="sxs-lookup"><span data-stu-id="48851-125">Using uppercase letters for the name of each custom function provides a consistent experience in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="48851-126">在 JSON 元数据文件中，确保每个 `id` 属性的值仅包含字母数字字符和句点。</span><span class="sxs-lookup"><span data-stu-id="48851-126">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="48851-127">在 JSON 元数据文件中，确保每个 `id` 属性的值在该文件范围内是唯一的。</span><span class="sxs-lookup"><span data-stu-id="48851-127">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="48851-128">也就是说，元数据文件中不应存在具有相同 `id` 值的两个函数对象。</span><span class="sxs-lookup"><span data-stu-id="48851-128">That is, no two function objects in the metadata file should have the same `id` value.</span></span> 

* <span data-ttu-id="48851-129">在将 JSON 元数据文件中的 `id` 属性的值与相应的 JavaScript 函数名称关联后，请勿再更改该值。</span><span class="sxs-lookup"><span data-stu-id="48851-129">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="48851-130">你可以通过更新 JSON 元数据文件中的 `name` 属性来更改最终用户在 Excel 中看到的函数名称，但绝不能更改已确定的 `id` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="48851-130">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="48851-131">在 JavaScript 文件中，请在同一位置指定所有自定义函数关联。</span><span class="sxs-lookup"><span data-stu-id="48851-131">In the JavaScript file, specify all custom function associations in the same location.</span></span> <span data-ttu-id="48851-132">例如，以下代码示例定义了两个自定义函数，并接着指定了这两个函数的关联信息。</span><span class="sxs-lookup"><span data-stu-id="48851-132">For example, the following code sample defines two custom functions and then specifies the association information for both functions.</span></span>

    ```js
    function add(first, second){
      return first + second;
    }

    function increment(incrementBy, callback) {
      var result = 0;
      var timer = setInterval(function() {
        result += incrementBy;
        callback.setResult(result);
      }, 1000);

      callback.onCanceled = function() {
        clearInterval(timer);
      };
    }

    // associate `id` values in the JSON metadata file to JavaScript function names
    CustomFunctions.associate("ADD", add);
    CustomFunctions.associate("INCREMENT", increment);
    ```

    <span data-ttu-id="48851-133">以下示例显示了与此 JavaScript 代码示例中定义的函数相对应的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="48851-133">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="48851-134">请注意，在此文件中，`id` 和 `name` 属性为大写字母。</span><span class="sxs-lookup"><span data-stu-id="48851-134">Note that the `id` and `name` properties are in uppercase letters in this file.</span></span> 

    ```json
    {
      "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
      "functions": [
        {
          "id": "ADD",
          "name": "ADD",
          ...
        },
        {
          "id": "INCREMENT",
          "name": "INCREMENT",
          ...
        }
      ]
    }
    ```

## <a name="declaring-optional-parameters"></a><span data-ttu-id="48851-135">声明可选参数</span><span class="sxs-lookup"><span data-stu-id="48851-135">Declaring optional parameters</span></span> 

<span data-ttu-id="48851-136">在 Excel for Windows（版本 1812 或更高版本）中，可以声明自定义函数的可选参数。</span><span class="sxs-lookup"><span data-stu-id="48851-136">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="48851-137">当用户在 Excel 中调用函数时，可选参数将显示在括号中。</span><span class="sxs-lookup"><span data-stu-id="48851-137">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="48851-138">例如，具有一个名为 `parameter1` 的必需参数和一个名为 `parameter2` 的可选参数的函数 `FOO` 将在 Excel 中显示为 `=FOO(parameter1, [parameter2])`。</span><span class="sxs-lookup"><span data-stu-id="48851-138">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="48851-139">若要使某个参数可选，请在定义函数的 JSON 元数据文件中将 `"optional": true` 添加到该参数。</span><span class="sxs-lookup"><span data-stu-id="48851-139">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="48851-140">以下示例显示对于函数 `=ADD(first, second, [third])` 会是怎么样的。</span><span class="sxs-lookup"><span data-stu-id="48851-140">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="48851-141">请注意，可选 `[third]` 参数后跟两个必需参数。</span><span class="sxs-lookup"><span data-stu-id="48851-141">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="48851-142">必需参数将先显示在 Excel 的公式 UI 中。</span><span class="sxs-lookup"><span data-stu-id="48851-142">Required parameters will appear first in Excel’s Formula UI.</span></span>

```json
{
    "id": "ADD",
    "name": "ADD",
    "description": "Add two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
        },
    "parameters": [
        {
            "name": "first",
            "description": "first number to add",
            "type": "number",
            "dimensionality": "scalar"
        },
        {
            "name": "second",
            "description": "second number to add",
            "type": "number",
            "dimensionality": "scalar",
        },
        {
            "name": "third",
            "description": "third optional number to add",
            "type": "number",
            "dimensionality": "scalar",
            "optional": true
        }
    ],
    "options": {
        "sync": false
    }
}
```

<span data-ttu-id="48851-143">定义包含一个或多个可选参数的函数时，应指定未定义可选参数时会发生什么情况。</span><span class="sxs-lookup"><span data-stu-id="48851-143">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="48851-144">在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。</span><span class="sxs-lookup"><span data-stu-id="48851-144">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="48851-145">如果未定义 `zipCode` 参数，则会将默认值设置为 98052。</span><span class="sxs-lookup"><span data-stu-id="48851-145">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="48851-146">如果未定义 `dayOfWeek` 参数，则会将其设置为星期三。</span><span class="sxs-lookup"><span data-stu-id="48851-146">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

```js
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek
  // ...
}
```

## <a name="additional-considerations"></a><span data-ttu-id="48851-147">其他注意事项</span><span class="sxs-lookup"><span data-stu-id="48851-147">Additional considerations</span></span>

<span data-ttu-id="48851-148">为创建一个可在多个平台（Office 外接程序的关键租户之一）上运行的外接程序，请勿访问自定义函数中的文档对象模型 (DOM) 或使用 jQuery 等这类依赖于 DOM 的库。</span><span class="sxs-lookup"><span data-stu-id="48851-148">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="48851-149">在自定义函数会使用 [JavaScript 运行时](custom-functions-runtime.md)的 Windows 版 Excel 中，自定义函数无法访问 DOM。</span><span class="sxs-lookup"><span data-stu-id="48851-149">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="48851-150">另请参阅</span><span class="sxs-lookup"><span data-stu-id="48851-150">See also</span></span>

* [<span data-ttu-id="48851-151">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="48851-151">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="48851-152">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="48851-152">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="48851-153">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="48851-153">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="48851-154">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="48851-154">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="48851-155">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="48851-155">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

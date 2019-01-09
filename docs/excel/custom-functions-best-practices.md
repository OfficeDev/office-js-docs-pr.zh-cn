---
ms.date: 01/08/2019
description: 了解在 Excel 中开发自定义函数的最佳实践。
title: 自定义函数最佳实践（预览）
ms.openlocfilehash: 45618a61d0d1fdd0398ecec3aa0db21e493787fd
ms.sourcegitcommit: 9afcb1bb295ec0c8940ed3a8364dbac08ef6b382
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2019
ms.locfileid: "27770649"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="07899-103">自定义函数最佳实践（预览）</span><span class="sxs-lookup"><span data-stu-id="07899-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="07899-104">本文介绍了在 Excel 中开发自定义函数的最佳实践。</span><span class="sxs-lookup"><span data-stu-id="07899-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="07899-105">错误处理</span><span class="sxs-lookup"><span data-stu-id="07899-105">Error handling</span></span>

<span data-ttu-id="07899-106">在生成定义自定义函数的外接程序时，请务必加入错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="07899-106">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="07899-107">自定义函数的错误处理与 [Excel JavaScript API 的错误处理](excel-add-ins-error-handling.md)大致相同。</span><span class="sxs-lookup"><span data-stu-id="07899-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="07899-108">在以下代码示例中，`.catch` 将处理之前发生在代码中的任何错误。</span><span class="sxs-lookup"><span data-stu-id="07899-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="troubleshooting"></a><span data-ttu-id="07899-109">故障排除</span><span class="sxs-lookup"><span data-stu-id="07899-109">Troubleshooting</span></span>

<span data-ttu-id="07899-110">如果要在 Windows 版 Office 中测试外接程序，则应启用**[运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)**，以解决外接程序的 XML 清单文件及多个安装和运行时条件问题。</span><span class="sxs-lookup"><span data-stu-id="07899-110">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="07899-111">运行时日志记录将 `console.log` 语句写入日志文件，以帮你发现问题。</span><span class="sxs-lookup"><span data-stu-id="07899-111">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

<span data-ttu-id="07899-112">若要向 Excel 自定义函数团队报告有关此故障排除方法的反馈，请发送团队反馈。</span><span class="sxs-lookup"><span data-stu-id="07899-112">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="07899-113">若要执行此操作，请选择“**文件|反馈|发送哭脸**”。</span><span class="sxs-lookup"><span data-stu-id="07899-113">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="07899-114">发送哭脸将提供必要的日志，以帮助我们了解你遇到的问题。</span><span class="sxs-lookup"><span data-stu-id="07899-114">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

## <a name="debugging"></a><span data-ttu-id="07899-115">调试</span><span class="sxs-lookup"><span data-stu-id="07899-115">Debugging</span></span>

<span data-ttu-id="07899-116">目前，有关调试 Excel 自定义函数的最佳方法是先[旁加载](../testing/sideload-office-add-ins-for-testing.md) **Excel Online** 内的外接程序。</span><span class="sxs-lookup"><span data-stu-id="07899-116">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="07899-117">然后，你可以通过使用[浏览器本机的 F12 调试工具](../testing/debug-add-ins-in-office-online.md)并结合以下技巧调试自定义函数：</span><span class="sxs-lookup"><span data-stu-id="07899-117">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="07899-118">使用自定义函数代码中的 `console.log` 语句，将输出实时发送到控制台。</span><span class="sxs-lookup"><span data-stu-id="07899-118">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="07899-119">使用自定义函数代码中的 `debugger;` 语句来指定当 F12 窗口打开时执行将暂停的断点。</span><span class="sxs-lookup"><span data-stu-id="07899-119">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="07899-120">例如，如果在 F12 窗口打开时运行以下函数，则执行将在 `debugger;` 语句上暂停，使你可以在函数返回之前手动检查参数值。</span><span class="sxs-lookup"><span data-stu-id="07899-120">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="07899-121">当 F12 窗口未打开时，`debugger;` 语句在 Excel Online 中无效。</span><span class="sxs-lookup"><span data-stu-id="07899-121">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="07899-122">目前，`debugger;` 语句对 Windows 版 Excel 无效。</span><span class="sxs-lookup"><span data-stu-id="07899-122">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="07899-123">如果你的外接程序无法注册，请验证是否为托管外接应用程序的 Web 服务器[正确配置了 SSL 证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="07899-123">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="07899-124">将函数名称与 JSON 元数据相关联</span><span class="sxs-lookup"><span data-stu-id="07899-124">Associating function names with JSON metadata</span></span>

<span data-ttu-id="07899-125">如[自定义函数概述](custom-functions-overview.md)文章中所述，自定义函数项目必须包含 JSON 元数据文件和脚本（JavaScript 或 TypeScript）文件才能构成完整的函数。</span><span class="sxs-lookup"><span data-stu-id="07899-125">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="07899-126">要使函数正常工作，需要将脚本文件中的函数名称绑定到 JSON 文件中列出的 ID。</span><span class="sxs-lookup"><span data-stu-id="07899-126">For a function to work properly, you'll need to bind the name of the function in the script file to the id listed in the JSON file.</span></span> <span data-ttu-id="07899-127">此过程称为“关联”。</span><span class="sxs-lookup"><span data-stu-id="07899-127">This process is called association.</span></span> <span data-ttu-id="07899-128">请记住在 JavaScript 代码文件的末尾包含关联，否则，函数将无法正常工作。</span><span class="sxs-lookup"><span data-stu-id="07899-128">Make a note to include associations at the end of your JavaScript code files; otherwise, your functions will not work.</span></span>

<span data-ttu-id="07899-129">以下代码示例展示了如何执行此关联操作。</span><span class="sxs-lookup"><span data-stu-id="07899-129">The following multi-part POST code shows how to do this.</span></span> <span data-ttu-id="07899-130">该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。</span><span class="sxs-lookup"><span data-stu-id="07899-130">For example, the following code sample defines the custom function  and then specifies that the function  corresponds to the object in the JSON metadata file where the value of the  property is ADD.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add); 
```

<span data-ttu-id="07899-131">在 JavaScript 文件中创建自定义函数和在 JSON 元数据文件中指定相应信息时，请记住以下最佳实践。</span><span class="sxs-lookup"><span data-stu-id="07899-131">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="07899-132">在 JSON 元数据文件中，函数的 `name` 和 `id` 只能使用大写字母。</span><span class="sxs-lookup"><span data-stu-id="07899-132">Only use uppercase letters for a function's `name` and `id` in the JSON metadata file.</span></span> <span data-ttu-id="07899-133">不要使用大小写字母混合或仅使用小写字母。</span><span class="sxs-lookup"><span data-stu-id="07899-133">Do not use a mix of cases or only lowercase letters.</span></span> <span data-ttu-id="07899-134">如果这样做，你最终可能会得到两个值，这些值只会因情况而异，从而导致意外覆盖你的函数。</span><span class="sxs-lookup"><span data-stu-id="07899-134">If you do, you may end up with two values that only differ by case which will cause unintentional overwriting of your functions.</span></span> <span data-ttu-id="07899-135">例如，`id` 值为 **add** 的函数对象稍后可以通过声明在 `id` 值为 **ADD** 的函数对象文件中覆盖。</span><span class="sxs-lookup"><span data-stu-id="07899-135">For example, a function object with an `id` value of **add** could be overwritten by declaration later in the file of function object with an `id` value of **ADD**.</span></span> <span data-ttu-id="07899-136">此外，`name` 属性还会定义最终用户将在 Excel 中看到的函数名称。</span><span class="sxs-lookup"><span data-stu-id="07899-136">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="07899-137">使用大写字母作为每个自定义函数的名称可在 Excel 中提供一致的体验，其中所有内置函数名称均为大写。</span><span class="sxs-lookup"><span data-stu-id="07899-137">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="07899-138">但是，在关联时没有必要将函数的 `name` 大写。</span><span class="sxs-lookup"><span data-stu-id="07899-138">However, it is not necessary to capitalize the function's `name` when associating.</span></span> <span data-ttu-id="07899-139">例如，`CustomFunctions.associate("add", add)` 等同于 `CustomFunctions.associate("ADD", add)`。</span><span class="sxs-lookup"><span data-stu-id="07899-139">For example, `CustomFunctions.associate("add", add)` is equivalent to `CustomFunctions.associate("ADD", add)`.</span></span>

* <span data-ttu-id="07899-140">在 JSON 元数据文件中，确保每个 `id` 属性的值仅包含字母数字字符和句点。</span><span class="sxs-lookup"><span data-stu-id="07899-140">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="07899-141">在 JSON 元数据文件中，确保每个 `id` 属性的值在该文件范围内是唯一的。</span><span class="sxs-lookup"><span data-stu-id="07899-141">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="07899-142">也就是说，元数据文件中不应存在具有相同 `id` 值的两个函数对象。</span><span class="sxs-lookup"><span data-stu-id="07899-142">That is, no two function objects in the metadata file should have the same `id` value.</span></span> 

* <span data-ttu-id="07899-143">在将 JSON 元数据文件中的 `id` 属性的值与相应的 JavaScript 函数名称关联后，请勿再更改该值。</span><span class="sxs-lookup"><span data-stu-id="07899-143">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="07899-144">你可以通过更新 JSON 元数据文件中的 `name` 属性来更改最终用户在 Excel 中看到的函数名称，但绝不能更改已确定的 `id` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="07899-144">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="07899-145">在 JavaScript 文件中，请在同一位置指定所有自定义函数关联。</span><span class="sxs-lookup"><span data-stu-id="07899-145">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="07899-146">例如，以下代码示例定义了两个自定义函数，并接着指定了这两个函数的关联信息。</span><span class="sxs-lookup"><span data-stu-id="07899-146">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="07899-147">以下示例显示了与此 JavaScript 代码示例中定义的函数相对应的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="07899-147">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="07899-148">请注意，在此文件中，`id` 和 `name` 属性为大写字母。</span><span class="sxs-lookup"><span data-stu-id="07899-148">Note that the `id` and `name` properties are in uppercase letters in this file.</span></span> 

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

## <a name="declaring-optional-parameters"></a><span data-ttu-id="07899-149">声明可选参数</span><span class="sxs-lookup"><span data-stu-id="07899-149">Declaring optional parameters</span></span> 
<span data-ttu-id="07899-150">在 Excel for Windows（版本 1812 或更高版本）中，可以声明自定义函数的可选参数。</span><span class="sxs-lookup"><span data-stu-id="07899-150">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="07899-151">当用户在 Excel 中调用函数时，可选参数将显示在括号中。</span><span class="sxs-lookup"><span data-stu-id="07899-151">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="07899-152">例如，具有一个名为 `parameter1` 的必需参数和一个名为 `parameter2` 的可选参数的函数 `FOO` 将在 Excel 中显示为 `=FOO(parameter1, [parameter2])`。</span><span class="sxs-lookup"><span data-stu-id="07899-152">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="07899-153">若要使某个参数可选，请在定义函数的 JSON 元数据文件中将 `"optional": true` 添加到该参数。</span><span class="sxs-lookup"><span data-stu-id="07899-153">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="07899-154">以下示例显示对于函数 `=ADD(first, second, [third])` 会是怎么样的。</span><span class="sxs-lookup"><span data-stu-id="07899-154">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="07899-155">请注意，可选 `[third]` 参数后跟两个必需参数。</span><span class="sxs-lookup"><span data-stu-id="07899-155">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="07899-156">必需参数将先显示在 Excel 的公式 UI 中。</span><span class="sxs-lookup"><span data-stu-id="07899-156">Required parameters will appear first in Excel’s Formula UI.</span></span>

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

<span data-ttu-id="07899-157">定义包含一个或多个可选参数的函数时，应指定未定义可选参数时会发生什么情况。</span><span class="sxs-lookup"><span data-stu-id="07899-157">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="07899-158">在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。</span><span class="sxs-lookup"><span data-stu-id="07899-158">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="07899-159">如果未定义 `zipCode` 参数，则会将默认值设置为 98052。</span><span class="sxs-lookup"><span data-stu-id="07899-159">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="07899-160">如果未定义 `dayOfWeek` 参数，则会将其设置为星期三。</span><span class="sxs-lookup"><span data-stu-id="07899-160">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="07899-161">其他注意事项</span><span class="sxs-lookup"><span data-stu-id="07899-161">Additional considerations</span></span>

<span data-ttu-id="07899-162">为创建一个可在多个平台（Office 外接程序的关键租户之一）上运行的外接程序，请勿访问自定义函数中的文档对象模型 (DOM) 或使用 jQuery 等这类依赖于 DOM 的库。</span><span class="sxs-lookup"><span data-stu-id="07899-162">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="07899-163">在自定义函数会使用 [JavaScript 运行时](custom-functions-runtime.md)的 Windows 版 Excel 中，自定义函数无法访问 DOM。</span><span class="sxs-lookup"><span data-stu-id="07899-163">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="07899-164">另请参阅</span><span class="sxs-lookup"><span data-stu-id="07899-164">See also</span></span>

* [<span data-ttu-id="07899-165">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="07899-165">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="07899-166">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="07899-166">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="07899-167">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="07899-167">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="07899-168">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="07899-168">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="07899-169">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="07899-169">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

---
ms.date: 10/17/2018
description: 了解 Excel 自定义函数的最佳做法和建议的模式。
title: 自定义函数最佳做法
ms.openlocfilehash: 10ba29966c1e991ca23674ce3e5da88de2772e00
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/19/2018
ms.locfileid: "25639999"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="e2b54-103">自定义函数的最佳做法（预览）</span><span class="sxs-lookup"><span data-stu-id="e2b54-103">Custom functions best practices</span></span>

<span data-ttu-id="e2b54-104">本文介绍在 Excel 中开发自定义函数的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="e2b54-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="e2b54-105">错误处理</span><span class="sxs-lookup"><span data-stu-id="e2b54-105">Error handling</span></span>

<span data-ttu-id="e2b54-p101">构建用来定义自定义函数的加载项时，请确保包含错误处理逻辑以解决运行时错误。自定义函数的错误处理与 [Excel JavaScript API 的错误处理](excel-add-ins-error-handling.md)  相同。 在以下代码示例中， `.catch` 将处理先前在代码中发生的任何错误。</span><span class="sxs-lookup"><span data-stu-id="e2b54-p101">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="debugging"></a><span data-ttu-id="e2b54-109">调试</span><span class="sxs-lookup"><span data-stu-id="e2b54-109">Debugging</span></span>

<span data-ttu-id="e2b54-p102">目前，调试 Excel 自定义函数的最佳方法是到首先在 **Excel Online** 内[旁加载](../testing/sideload-office-add-ins-for-testing.md)加载项。然后，可以通过结合下列方法使用[浏览器自带的 F12 键调试工具](../testing/debug-add-ins-in-office-online.md)来调试自定义函数：</span><span class="sxs-lookup"><span data-stu-id="e2b54-p102">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="e2b54-112">使用自定义函数代码中的 `console.log` 语句发送输出到实时控制台。</span><span class="sxs-lookup"><span data-stu-id="e2b54-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="e2b54-p103">使用自定义函数代码中的 `debugger;` 语句来指定当 F12 窗口打开时执行将暂停的断点。例如，如果 F12 窗口处于打开状态时以下函数正在运行，执行将暂停于 `debugger;` 语句上，使您能够在函数返回前手动检查参数值。当 F12 窗口未打开时， `debugger;` 语句在 Excel Online 中就无效。目前， `debugger;` 语句在 Excel for Windows 中无效。</span><span class="sxs-lookup"><span data-stu-id="e2b54-p103">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open. For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns. The `debugger;` statement has no effect in Excel Online when the F12 window is not open. Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="e2b54-117">如果加载项无法注册，请 [验证为托管加载项应用程序的 Web 服务器正确配置了 SSL 证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) 。</span><span class="sxs-lookup"><span data-stu-id="e2b54-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="e2b54-118">如果您在 Windows 桌面中测试 Office 中的加载项，可以启用 [运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) ，以调试加载项 XML 清单文件以及若干安装和运行时条件等问题。</span><span class="sxs-lookup"><span data-stu-id="e2b54-118">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="e2b54-119">函数名称映射到 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="e2b54-119">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="e2b54-p104">如 [自定义函数概述](custom-functions-overview.md) 文章中所述，自定义函数项目必须包含一个 JSON 元数据文件，它提供了 Excel 要求注册自定义函数，并可提供给最终用户的信息。此外，要在 JavaScript 文件中定义您的自定义函数，必须提供信息以指定，JSON 元数据文件中的哪一个函数对象对应于每个 JavaScript 文件中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="e2b54-p104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="e2b54-122">例如，下面的代码示例定义了自定义函数 `add` ，然后指定函数 `add` 对应于 JSON 元数据文件中 `id` 属性值是 **ADD** 的对象。</span><span class="sxs-lookup"><span data-stu-id="e2b54-122">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="e2b54-123">在 JavaScript 文件中创建自定义函数并指定 JSON 元数据文件中的对应信息时，请记住以下最佳做法。</span><span class="sxs-lookup"><span data-stu-id="e2b54-123">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="e2b54-p105">在 JavaScript 文件中，指定 camelCase 的函数名称。例如，函数名称 `addTenToInput` 在 camelCase 中编写为：名称的第一个单词以小写字母开头，名称中的每个后续单词以大写字母开头。</span><span class="sxs-lookup"><span data-stu-id="e2b54-p105">In the JavaScript file, specify function names in camelCase. For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="e2b54-p106">在 JSON 元数据文件中，以大写形式指定每个 `name` 属性的值。 `name` 属性定义了最终用户将在 Excel 中看到的函数名称。为每个自定义的函数名称使用大写字母可以为最终用户在 Excel 中提供一致的体验，在那里所有内置的函数名称都是大写的。</span><span class="sxs-lookup"><span data-stu-id="e2b54-p106">In the JSON metadata file, specify the value of each `name` property in uppercase. The `name` property defines the function name that end users will see in Excel. Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="e2b54-p107">在 JSON 元数据文件中，以大写形式指定每个 `id` 属性的值。这样做可以明显地知晓，您的 JavaScript 代码中的 `CustomFunctionMappings` 语句的哪一部分对应于 JSON 元数据文件的  `id` 属性（前提是您的函数名称如前面所建议的使用 camelCase）。</span><span class="sxs-lookup"><span data-stu-id="e2b54-p107">In the JSON metadata file, specify the value of each `id` property in uppercase. Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="e2b54-131">在 JSON 元数据文件中，确保每个 `id` 属性的值仅包含字母数字字符和句号。</span><span class="sxs-lookup"><span data-stu-id="e2b54-131">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span> 

* <span data-ttu-id="e2b54-p108">在 JSON 元数据文件中，确保每个 `id` 属性的值在此文件范围内是唯一的。即元数据文件中不存在两个函数对象同时具有相同的 `id` 值。此外，在元数据文件中不指定两个仅仅是大小写不同的 `id` 值。例如，不会定义一个具有 **add** 的 `id` 值的函数对象和一个具有 **ADD** 的 `id` 值的函数对象。</span><span class="sxs-lookup"><span data-stu-id="e2b54-p108">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file. That is, no two function objects in the metadata file should have the same `id` value. Additionally, do not specify two `id` values in the metadata file that only differ by case. For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="e2b54-p109">在 JSON 元数据文件被映射到相应的 JavaScript 函数名称后，不要更改其中的 `id` 属性的值。您可以通过在 JSON元数据文件中更新 `name` 属性来更改最终用户可在 Excel 中看到的函数名称，但在 `id` 属性被建立后，您永远都不应更改它的值。</span><span class="sxs-lookup"><span data-stu-id="e2b54-p109">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name. You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="e2b54-p110">在 JavaScript 文件中，在相同位置指定所有自定义函数的映射。例如，下面的代码示例定义了两个自定义的函数，并指定这两个函数的映射信息。</span><span class="sxs-lookup"><span data-stu-id="e2b54-p110">In the JavaScript file, specify all custom function mappings in the same location. For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    // map `id` values in the JSON metadata file to JavaScript function names
    CustomFunctionMappings.ADD = add;
    CustomFunctionMappings.INCREMENT = increment;
    ```

    <span data-ttu-id="e2b54-140">下面的示例演示对应于此 JavaScript 代码示例中定义的函数的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="e2b54-140">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="e2b54-141">其他注意事项</span><span class="sxs-lookup"><span data-stu-id="e2b54-141">Additional considerations</span></span>

<span data-ttu-id="e2b54-142">为了创建一个将在多个平台（Office 加载项的关键租户之一）上运行的加载项，您不应该访问自定义函数中的  Document Object Model（DOM）或使用像 jQuery 那样依赖于 DOM 的库。</span><span class="sxs-lookup"><span data-stu-id="e2b54-142">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="e2b54-143">在 Excel for Windows 中，自定义函数使用 [JavaScript 运行时](custom-functions-runtime.md) ，自定义函数无法访问 DOM。</span><span class="sxs-lookup"><span data-stu-id="e2b54-143">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="e2b54-144">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e2b54-144">See also</span></span>

* [<span data-ttu-id="e2b54-145">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="e2b54-145">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="e2b54-146">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="e2b54-146">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="e2b54-147">Excel 自定义函数运行时</span><span class="sxs-lookup"><span data-stu-id="e2b54-147">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="e2b54-148">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="e2b54-148">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

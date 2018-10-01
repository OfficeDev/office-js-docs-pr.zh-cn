---
ms.date: 09/27/2018
description: 了解 Excel 自定义函数的最佳做法和建议的模式。
title: 自定义函数的最佳做法
ms.openlocfilehash: d157464a3a8bf453cd0970281f1a4fdd27df5d25
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348785"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="18a43-103">自定义函数的最佳做法（预览）</span><span class="sxs-lookup"><span data-stu-id="18a43-103">Custom functions best practices</span></span>

<span data-ttu-id="18a43-104">本文介绍在 Excel 中开发自定义函数的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="18a43-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="18a43-105">错误处理</span><span class="sxs-lookup"><span data-stu-id="18a43-105">Error handling</span></span>

<span data-ttu-id="18a43-106">构建定义自定义函数的加载项时，请务必加入错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="18a43-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="18a43-107">自定义函数的错误处理与 [Excel JavaScript API 的错误处理大体相同](excel-add-ins-error-handling.md)。</span><span class="sxs-lookup"><span data-stu-id="18a43-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="18a43-108">在下面的代码示例中，`.catch` 将处理之前出现在代码中的任何错误。</span><span class="sxs-lookup"><span data-stu-id="18a43-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="debugging"></a><span data-ttu-id="18a43-109">调试</span><span class="sxs-lookup"><span data-stu-id="18a43-109">Debugging</span></span>

<span data-ttu-id="18a43-p102">目前，调试 Excel 自定义函数的最佳方法是到首先在 **Excel Online** 内[旁加载](../testing/sideload-office-add-ins-for-testing.md)加载项。然后，可以通过结合下列方法使用[浏览器自带的 F12 键调试工具](../testing/debug-add-ins-in-office-online.md)来调试自定义函数：</span><span class="sxs-lookup"><span data-stu-id="18a43-p102">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md). Use  statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="18a43-112">使用自定义函数代码中的 `console.log` 语句实时发送输出到控制台。</span><span class="sxs-lookup"><span data-stu-id="18a43-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="18a43-113">使用自定义函数代码内的 `debugger;` 语句来指定 F12 窗口打开时执行将暂停的断点。</span><span class="sxs-lookup"><span data-stu-id="18a43-113">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="18a43-114">例如，如果 F12 窗口打开时以下函数运行，则执行将在 `debugger;` 语句上暂停，使你能够在函数返回前手动检查参数值。</span><span class="sxs-lookup"><span data-stu-id="18a43-114">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="18a43-115">`debugger;` 语句在 F12 窗口未打开时在 Excel Online 中不起作用。</span><span class="sxs-lookup"><span data-stu-id="18a43-115">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="18a43-116">目前，`debugger;` 语句在 Excel for Windows 中不起作用。</span><span class="sxs-lookup"><span data-stu-id="18a43-116">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="18a43-117">如果加载项无法注册，请[验证为托管加载项应用程序的 Web 服务器正确配置了 SSL 证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="18a43-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="18a43-118">如果您在 Office 2016 桌面版中测试加载项，可以启用[运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)，以调试加载项 XML 指令清单文件以及若干安装和运行时条件相关的问题。</span><span class="sxs-lookup"><span data-stu-id="18a43-118">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="18a43-119">函数名称映射到 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="18a43-119">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="18a43-120">如[自定义函数概述](custom-functions-overview.md)文章中所述，自定义函数项目必须包含一个 JSON 元数据文件，它提供了 Excel 注册并向最终用户提供自定义函数所需的信息。</span><span class="sxs-lookup"><span data-stu-id="18a43-120">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="18a43-121">此外，在定义了自定义函数的 JavaScript 文件中，必须提供信息以指定 JavaScript 文件中的每个自定义函数对应于 JSON 元数据文件中的哪些函数对象。</span><span class="sxs-lookup"><span data-stu-id="18a43-121">Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="18a43-122">例如，下面的代码示例定义了自定义函数 `add`，然后指定函数 `add` 对应于 JSON 元数据文件中 `id` 属性值是 **ADD** 的对象。</span><span class="sxs-lookup"><span data-stu-id="18a43-122">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="18a43-123">在 JavaScript 文件中创建自定义函数并指定 JSON 元数据文件中的对应信息时，请记住以下最佳做法。</span><span class="sxs-lookup"><span data-stu-id="18a43-123">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="18a43-124">在 JavaScript 文件中，以驼峰拼写法指定函数名称。</span><span class="sxs-lookup"><span data-stu-id="18a43-124">In the JavaScript file, specify function names in camelCase.</span></span> <span data-ttu-id="18a43-125">例如，函数名称 `addTenToInput` 使用了驼峰拼写法：名称的第一个单词开头为小写字母，名称中的每个后续单词开头为大写字母。</span><span class="sxs-lookup"><span data-stu-id="18a43-125">For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="18a43-126">在 JSON 元数据文件中，以大写字母指定每个 `name` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="18a43-126">In the JSON metadata file, specify the value of each `name` property in uppercase.</span></span> <span data-ttu-id="18a43-127">`name` 属性定义了最终用户将在 Excel 中看到的函数名称。</span><span class="sxs-lookup"><span data-stu-id="18a43-127">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="18a43-128">在每个自定义函数名称中使用大写字母为最终用户提供了一致的体验，因为在 Excel 中所有内置函数名称都是大写。</span><span class="sxs-lookup"><span data-stu-id="18a43-128">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="18a43-129">在 JSON 元数据文件中，以大写字母指定每个 `id` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="18a43-129">In the JSON metadata file, specify the value of each `id` property in uppercase.</span></span> <span data-ttu-id="18a43-130">这样使 `CustomFunctionMappings` JavaScript 代码中对应于 JSON 元数据文件 `id` 属性的部分显而易见（前提是函数名称使用驼峰拼写法，如前面所建议）。</span><span class="sxs-lookup"><span data-stu-id="18a43-130">Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="18a43-131">在 JSON 元数据文件中，确保每个 `id` 属性的值在文件范围内是唯一的。</span><span class="sxs-lookup"><span data-stu-id="18a43-131">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="18a43-132">即，元数据文件中没有两个函数对象具有相同的 `id` 值。</span><span class="sxs-lookup"><span data-stu-id="18a43-132">That is, no two function objects in the metadata file should have the same `id` value.</span></span> <span data-ttu-id="18a43-133">此外，不要在元数据文件中指定两个 `id` 仅大小写不同的值。</span><span class="sxs-lookup"><span data-stu-id="18a43-133">Additionally, do not specify two `id` values in the metadata file that only differ by case.</span></span> <span data-ttu-id="18a43-134">例如，不要将一个函数对象的 `id` 值定义为 **add**，而将另一函数对象的 `id` 值定义为 \*\* ADD\*\*。</span><span class="sxs-lookup"><span data-stu-id="18a43-134">For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="18a43-135">JSON 元数据文件中的 `id` 属性值在映射到相应的 JavaScript 函数名称后，不要更改其值。</span><span class="sxs-lookup"><span data-stu-id="18a43-135">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="18a43-136">您可以通过更新 JSON 元数据文件中 `name` 属性的值更改用户在 Excel 中看到的函数名称，但永远不要在 `id` 属性已建立后更改其值。</span><span class="sxs-lookup"><span data-stu-id="18a43-136">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="18a43-137">在 JavaScript 文件中，在相同位置指定所有自定义函数映射。</span><span class="sxs-lookup"><span data-stu-id="18a43-137">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="18a43-138">例如，下面的代码示例定义两个自定义的函数，并指定这两个函数的映射信息。</span><span class="sxs-lookup"><span data-stu-id="18a43-138">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="18a43-139">下面的示例演示对应于此 JavaScript 代码示例中定义的函数的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="18a43-139">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="18a43-140">另请参阅</span><span class="sxs-lookup"><span data-stu-id="18a43-140">See also</span></span>

* [<span data-ttu-id="18a43-141">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="18a43-141">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="18a43-142">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="18a43-142">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="18a43-143">Excel 自定义函数运行时</span><span class="sxs-lookup"><span data-stu-id="18a43-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="18a43-144">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="18a43-144">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

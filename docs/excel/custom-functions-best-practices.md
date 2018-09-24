---
ms.date: 09/20/2018
description: 了解 Excel 自定义函数的最佳实践和建议的模式。
title: 自定义函数的最佳实践
ms.openlocfilehash: 1f2c0a80e62b65523fcc1673ba2ca4be444e6ce0
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/21/2018
ms.locfileid: "24068811"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="d9f8a-103">自定义函数的最佳实践</span><span class="sxs-lookup"><span data-stu-id="d9f8a-103">Custom functions best practices</span></span>

<span data-ttu-id="d9f8a-104">本文介绍在 Excel 中开发自定义函数的最佳实践。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="d9f8a-105">错误处理</span><span class="sxs-lookup"><span data-stu-id="d9f8a-105">Error handling</span></span>

<span data-ttu-id="d9f8a-106">构建定义了自定义函数的加载项时，请务必加入错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="d9f8a-107">自定义的函数的错误处理与 [Excel JavaScript API 的错误处理整体类同](excel-add-ins-error-handling.md)。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="d9f8a-108">在下面的代码示例中，`.catch` 将处理之前出现在代码中的任何错误。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    let url = "https://jsonplaceholder.typicode.com/comments/" + x; 
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

## <a name="error-logging"></a><span data-ttu-id="d9f8a-109">错误日志记录</span><span class="sxs-lookup"><span data-stu-id="d9f8a-109">Error logging</span></span>

<span data-ttu-id="d9f8a-110">有多种方式启用自定义函数加载项的用错误日志记录，例如：</span><span class="sxs-lookup"><span data-stu-id="d9f8a-110">You can enable error logging for your custom functions add-in in multiple ways, such as:</span></span> 

- <span data-ttu-id="d9f8a-111">[使用运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest)调试外接程序的 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-111">[Use runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) to debug your add-in's XML manifest file.</span></span> 

- <span data-ttu-id="d9f8a-112">使用自定义函数代码中的 `console.log` 语句发送输出到实时控制台。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

> [!NOTE]
> <span data-ttu-id="d9f8a-113">运行时日志记录功能目前仅适用于 Office 2016 桌面版。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-113">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

## <a name="debugging"></a><span data-ttu-id="d9f8a-114">调试</span><span class="sxs-lookup"><span data-stu-id="d9f8a-114">Debugging</span></span>

<span data-ttu-id="d9f8a-115">目前，调试 Excel 自定义函数的最佳方法是使用 [Excel Online](https://www.office.com/launch/excel)，并使用浏览器内部的 f12 键调试工具。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-115">Currently, the best method for debugging Excel custom functions is to use [Excel Online](https://www.office.com/launch/excel) and use the F12 debugging tool native to your browser.</span></span> <span data-ttu-id="d9f8a-116">将来可能存在自定义函数的其他调试工具。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-116">Additional debugging tools for custom functions may be available in the future.</span></span>

## <a name="mapping-names"></a><span data-ttu-id="d9f8a-117">映射名称</span><span class="sxs-lookup"><span data-stu-id="d9f8a-117">Mapping names</span></span>

<span data-ttu-id="d9f8a-118">默认情况下，JavaScript 文件中的自定义函数的名称通常完全使用大写字母声明，并与用户在 Excel 中看到的函数名称相对应。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-118">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="d9f8a-119">但是，可以通过使用 `CustomFunctionsMappings` 对象将一个或多个函数名称从 JavaScript 文件映射到不同值来更改此设置，用户将在 Excel 中看到的函数名称对应于这些值。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-119">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="d9f8a-120">虽然不要求您使用 `CustomFunctionsMapping`，但在使用难以处理大写函数名的 uglifier、webpack 或 import 语法时它可能非常有帮助。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-120">Although you're not required to use `CustomFunctionsMapping`, it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span>
  
<span data-ttu-id="d9f8a-121">下面的代码示例定义一个“键-值”对，将 JavaScript 函数名称 `plusFortyTwo` 映射到 Excel UI 中的 `ADD42` 函数名称。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-121">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="d9f8a-122">当用户选择 Excel 中的 `ADD42` 函数时，JavaScript 函数 `plusFortyTwo` 将运行。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-122">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="d9f8a-123">下面的代码示例定义两个“键-值”对。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-123">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="d9f8a-124">第一对将 JavaScript 函数名称 `plusFifty` 映射到 Excel UI 中的 `ADD50` 函数名称，第二对将 JavaScript 函数名称 `plusOneHundred` 映射到 Excel UI 中的 `ADD100` 函数名称。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-124">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="d9f8a-125">当用户选择 Excel 中的 `ADD50` 函数时，JavaScript 函数 `plusFifty` 将运行。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-125">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="d9f8a-126">当用户选择 Excel 中的 `ADD100` 函数时，JavaScript 函数 `plusOneHundred` 将运行。</span><span class="sxs-lookup"><span data-stu-id="d9f8a-126">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

```js
function plusFifty(num) {
    return num + 50;  
} 

function plusOneHundred(num) {
    return num + 100;  
}  
  
CustomFunctionsMappings = {
    "plusFifty" : ADD50,  
    "plusOneHundred" : ADD100
}
 ```

 ## <a name="see-also"></a><span data-ttu-id="d9f8a-127">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d9f8a-127">See also</span></span>

* [<span data-ttu-id="d9f8a-128">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="d9f8a-128">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="d9f8a-129">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="d9f8a-129">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d9f8a-130">Excel 自定义函数运行时</span><span class="sxs-lookup"><span data-stu-id="d9f8a-130">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
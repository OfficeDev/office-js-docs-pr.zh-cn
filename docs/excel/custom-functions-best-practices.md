---
ms.date: 09/20/2018
description: 了解 Excel 自定义函数的最佳实践和建议的模式。
title: 自定义函数的最佳实践
ms.openlocfilehash: 4fe0ddc36ce1b08ea360bb556121e76cd57c3823
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004908"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="3ca97-103">自定义函数的最佳实践</span><span class="sxs-lookup"><span data-stu-id="3ca97-103">Custom functions best practices</span></span>

<span data-ttu-id="3ca97-104">本文介绍在 Excel 中开发自定义函数的最佳实践。</span><span class="sxs-lookup"><span data-stu-id="3ca97-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="3ca97-105">错误处理</span><span class="sxs-lookup"><span data-stu-id="3ca97-105">Error handling</span></span>

<span data-ttu-id="3ca97-106">构建定义了自定义函数的加载项时，请务必加入错误处理逻辑，以便处理运行时错误。</span><span class="sxs-lookup"><span data-stu-id="3ca97-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="3ca97-107">自定义的函数的错误处理与 [Excel JavaScript API 的错误处理整体类同](excel-add-ins-error-handling.md)。</span><span class="sxs-lookup"><span data-stu-id="3ca97-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="3ca97-108">在下面的代码示例中，`.catch` 将处理之前出现在代码中的任何错误。</span><span class="sxs-lookup"><span data-stu-id="3ca97-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi.com/comments/" + x; 
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

## <a name="debugging"></a><span data-ttu-id="3ca97-109">调试</span><span class="sxs-lookup"><span data-stu-id="3ca97-109">Debugging</span></span>
<span data-ttu-id="3ca97-110">目前，调试 Excel 自定义函数的最佳方法是首先在 **Excel Online** 中[旁加载](../testing/sideload-office-add-ins-for-testing.md)加载项。</span><span class="sxs-lookup"><span data-stu-id="3ca97-110">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="3ca97-111">然后，您可以[使用浏览器内部的 F12 调试工具](../testing/debug-add-ins-in-office-online.md)来调试自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3ca97-111">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md).</span></span> <span data-ttu-id="3ca97-112">使用自定义函数代码中的 `console.log` 语句发送输出到实时控制台。</span><span class="sxs-lookup"><span data-stu-id="3ca97-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

<span data-ttu-id="3ca97-113">如果加载项无法注册，请[验证为托管加载项应用程序的 Web 服务器正确配置了 SSL 证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="3ca97-113">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="3ca97-114">如果您在 Office 2016 桌面中测试加载项，您可以启用[运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)，以调试加载项 XML 指令清单文件以及若干安装和运行时条件的问题。</span><span class="sxs-lookup"><span data-stu-id="3ca97-114">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span> 


## <a name="mapping-names"></a><span data-ttu-id="3ca97-115">映射名称</span><span class="sxs-lookup"><span data-stu-id="3ca97-115">Mapping names</span></span>

<span data-ttu-id="3ca97-116">默认情况下，JavaScript 文件中的自定义函数的名称通常完全使用大写字母声明，并与用户在 Excel 中看到的函数名称相对应。</span><span class="sxs-lookup"><span data-stu-id="3ca97-116">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="3ca97-117">但是，可以通过使用 `CustomFunctionsMappings` 对象将一个或多个函数名称从 JavaScript 文件映射到不同值来更改此设置，用户将在 Excel 中看到的函数名称对应于这些值。</span><span class="sxs-lookup"><span data-stu-id="3ca97-117">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="3ca97-118">如果您使用难以处理大写函数名称的 uglifier、webpack 或 import 语法，这非常有用。</span><span class="sxs-lookup"><span data-stu-id="3ca97-118">Although you're not required to use , it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span> <span data-ttu-id="3ca97-119">`CustomFunctionsMappings` 对使用 JavaScript 的项目可能是可选的，但如果您的项目使用 TypeScript， 则必须使用它。</span><span class="sxs-lookup"><span data-stu-id="3ca97-119">`CustomFunctionsMappings` is possibly optional for projects using JavaScript but must be used if your project uses TypeScript.</span></span>  
  
<span data-ttu-id="3ca97-120">下面的代码示例定义一个“键-值”对，将 JavaScript 函数名称 `plusFortyTwo` 映射到 Excel UI 中的 `ADD42` 函数名称。</span><span class="sxs-lookup"><span data-stu-id="3ca97-120">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="3ca97-121">当用户选择 Excel 中的 `ADD42` 函数时，JavaScript 函数 `plusFortyTwo` 将运行。</span><span class="sxs-lookup"><span data-stu-id="3ca97-121">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="3ca97-122">下面的代码示例定义两个“键-值”对。</span><span class="sxs-lookup"><span data-stu-id="3ca97-122">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="3ca97-123">第一对将 JavaScript 函数名称 `plusFifty` 映射到 Excel UI 中的 `ADD50` 函数名称，第二对将 JavaScript 函数名称 `plusOneHundred` 映射到 Excel UI 中的 `ADD100` 函数名称。</span><span class="sxs-lookup"><span data-stu-id="3ca97-123">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="3ca97-124">当用户选择 Excel 中的 `ADD50` 函数时，JavaScript 函数 `plusFifty` 将运行。</span><span class="sxs-lookup"><span data-stu-id="3ca97-124">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="3ca97-125">当用户选择 Excel 中的 `ADD100` 函数时，JavaScript 函数 `plusOneHundred` 将运行。</span><span class="sxs-lookup"><span data-stu-id="3ca97-125">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

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

 ## <a name="see-also"></a><span data-ttu-id="3ca97-126">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3ca97-126">See also</span></span>

- [<span data-ttu-id="3ca97-127">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="3ca97-127">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="3ca97-128">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="3ca97-128">Custom functions metadata</span></span>](custom-functions-json.md)
- [<span data-ttu-id="3ca97-129">Excel 自定义函数运行时</span><span class="sxs-lookup"><span data-stu-id="3ca97-129">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)

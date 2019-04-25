---
ms.date: 03/20/2019
description: 了解 Excel 自定义函数的运行时。
title: 自定义函数体系结构（预览版）
localization_priority: Priority
ms.openlocfilehash: b3f3d6c5eda51639a734c6d0f162c596f0c1e41b
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448600"
---
# <a name="custom-functions-architecture"></a><span data-ttu-id="39fe3-103">自定义函数体系结构</span><span class="sxs-lookup"><span data-stu-id="39fe3-103">Custom functions architecture</span></span>

 <span data-ttu-id="39fe3-104">自定义函数具有自己独特的运行时，可以优先执行计算。</span><span class="sxs-lookup"><span data-stu-id="39fe3-104">Custom functions are with their own unique runtime that prioritizes execution of calculations.</span></span> <span data-ttu-id="39fe3-105">本文将介绍自定义函数运行时与基于浏览器的 JavaScript 引擎之间的差异，该引擎支持加载项的其他绝大部分。</span><span class="sxs-lookup"><span data-stu-id="39fe3-105">This article will cover the differences between the custom functions runtime and the browser-based JavaScript engine which powers most other parts of your add-in.</span></span>

## <a name="custom-functions-runtime"></a><span data-ttu-id="39fe3-106">自定义函数运行时</span><span class="sxs-lookup"><span data-stu-id="39fe3-106">Custom functions runtime</span></span>

<span data-ttu-id="39fe3-107">Office Web 加载项可以作为任务窗格或内容窗格与用户进行交互，并且可以包括命令和自定义函数。</span><span class="sxs-lookup"><span data-stu-id="39fe3-107">An Office Web Add-in can interact with the user as a task pane, or a content pane, and can include commands and custom functions.</span></span> <span data-ttu-id="39fe3-108">所有这些部分都在浏览器引擎运行时中运行，自定义函数除外。</span><span class="sxs-lookup"><span data-stu-id="39fe3-108">All of these parts run in a browser engine runtime except for custom functions.</span></span> <span data-ttu-id="39fe3-109">自定义函数在单独的自定义函数运行时中运行，以优化计算速度。</span><span class="sxs-lookup"><span data-stu-id="39fe3-109">Custom functions run in a separate custom functions runtime to optimize for calculation speed.</span></span>

<span data-ttu-id="39fe3-110">请注意，如果你使用 [Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来生成项目，则自定义函数运行时将通过 functions.html 文件中引用的 custom-functions.js 脚本文件加载。</span><span class="sxs-lookup"><span data-stu-id="39fe3-110">Note that if you're using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to generate your project, the custom functions runtime will load through the custom-functions.js script file referenced in the functions.html file.</span></span> <span data-ttu-id="39fe3-111">functions.html 仅用于加载运行时，且不应用作加载项的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="39fe3-111">The functions.html serves only to load the runtime and shouldn't be used as the task pane for your add-in.</span></span>

<span data-ttu-id="39fe3-112">下表突出显示了自定义函数运行时与浏览器引擎运行时之间的差异：</span><span class="sxs-lookup"><span data-stu-id="39fe3-112">The following table highlights the differences between the custom functions runtime and the browser engine runtime:</span></span>

| <span data-ttu-id="39fe3-113">自定义函数运行时</span><span class="sxs-lookup"><span data-stu-id="39fe3-113">Custom functions runtime</span></span>  | <span data-ttu-id="39fe3-114">浏览器引擎运行时</span><span class="sxs-lookup"><span data-stu-id="39fe3-114">Browser engine runtime</span></span>    |
|------------------------------------------------------------------ |-------------------------------------------------------------------------------------------------------------- |
| <span data-ttu-id="39fe3-115">支持从单元格中返回值</span><span class="sxs-lookup"><span data-stu-id="39fe3-115">Supports returning a value from a cell</span></span>    | <span data-ttu-id="39fe3-116">支持 Office.js API 和 UI 元素</span><span class="sxs-lookup"><span data-stu-id="39fe3-116">Supports Office.js APIs and UI elements</span></span>   |
| <span data-ttu-id="39fe3-117">没有 `localStorage` 对象，改用 `AsyncStorage`</span><span class="sxs-lookup"><span data-stu-id="39fe3-117">Does not have `localStorage` object, instead uses `AsyncStorage`</span></span>  | <span data-ttu-id="39fe3-118">具有 `localStorage` 对象，可以选择使用 `AsyncStorage` 对象</span><span class="sxs-lookup"><span data-stu-id="39fe3-118">Has `localStorage` object, can optionally use `AsyncStorage` object</span></span>   |
| <span data-ttu-id="39fe3-119">不支持与 DOM 交互，或者加载依赖于 DOM 的库，如 jQuery。</span><span class="sxs-lookup"><span data-stu-id="39fe3-119">Does not support interacting with the DOM, or loading libraries that depend on the DOM such as jQuery.</span></span>    | <span data-ttu-id="39fe3-120">支持与 DOM 交互，和加载依赖于 DOM 的库。</span><span class="sxs-lookup"><span data-stu-id="39fe3-120">Supports interacting with the DOM and loading libraries that depend on the DOM.</span></span> |


## <a name="browser-engine-runtime"></a><span data-ttu-id="39fe3-121">浏览器引擎运行时</span><span class="sxs-lookup"><span data-stu-id="39fe3-121">Browser engine runtime</span></span>

<span data-ttu-id="39fe3-122">任务窗格、内容加载项和命令在浏览器引擎运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="39fe3-122">The task pane, content add-in, and commands run in a browser engine runtime.</span></span>

<span data-ttu-id="39fe3-123">浏览器引擎运行时支持 Office.js API。</span><span class="sxs-lookup"><span data-stu-id="39fe3-123">The browser engine runtime supports the Office.js APIs.</span></span> <span data-ttu-id="39fe3-124">请记住，任何 Excel API（例如允许你操作 Excel 表的 API）都可以在浏览器引擎运行时上运行，但无法从自定义函数运行时直接访问。</span><span class="sxs-lookup"><span data-stu-id="39fe3-124">Keep in mind that any of the Excel APIs, such as APIs which allow you to manipulate Excel tables, run on the browser engine runtime, but aren't directly accessible from the custom functions runtime.</span></span>

## <a name="communicate-between-runtimes"></a><span data-ttu-id="39fe3-125">运行时之间的通信</span><span class="sxs-lookup"><span data-stu-id="39fe3-125">Communicate between runtimes</span></span>

<span data-ttu-id="39fe3-126">你的自定义函数代码无法直接与 Web 加载项的其他部分（例如任务窗格）中的代码进行交互，因为它们位于不同的运行时。</span><span class="sxs-lookup"><span data-stu-id="39fe3-126">Your custom functions code cannot directly interact with code in other parts of your web add-in, like the task pane because they are in different runtimes.</span></span> <span data-ttu-id="39fe3-127">但在某些方案中，可能需要共享数据，例如传递令牌。</span><span class="sxs-lookup"><span data-stu-id="39fe3-127">But in some scenarios you may need to share data, such as passing a token.</span></span>

<span data-ttu-id="39fe3-128">`AsyncStorage` 可用于存储自定义函数的数据并从任务窗格代码中获取数据。</span><span class="sxs-lookup"><span data-stu-id="39fe3-128">`AsyncStorage` can be used to store data from your custom functions and get data from your task pane code.</span></span> <span data-ttu-id="39fe3-129">有关存储和共享数据的详细信息，请参阅[保存和共享状态](custom-functions-overview.md#saving-and-sharing-state)。</span><span class="sxs-lookup"><span data-stu-id="39fe3-129">For more information about storing and sharing data, see [Saving and sharing state](custom-functions-overview.md#saving-and-sharing-state).</span></span>

<span data-ttu-id="39fe3-130">可以使用这一专用于模式和做法的 [Github 存储库](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)中的 `AsyncStorage` 查看代码示例。</span><span class="sxs-lookup"><span data-stu-id="39fe3-130">You can see a code sample using `AsyncStorage` in this [Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) dedicated to patterns and practices.</span></span>
<span data-ttu-id="39fe3-131">有关 `AsyncStorage` 的更多常规信息，请参阅[自定义函数运行时](./custom-functions-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="39fe3-131">For more general information about `AsyncStorage`, see [Custom functions runtime](./custom-functions-runtime.md).</span></span>

<span data-ttu-id="39fe3-132">`AsyncStorage` 也可用于身份验证。</span><span class="sxs-lookup"><span data-stu-id="39fe3-132">`AsyncStorage` can also be useful for authentication.</span></span> <span data-ttu-id="39fe3-133">有关详细信息，请参阅[自定义函数身份验证](custom-functions-authentication.md)。</span><span class="sxs-lookup"><span data-stu-id="39fe3-133">For more information, see [Custom functions authentication](custom-functions-authentication.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="39fe3-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="39fe3-134">See also</span></span>

* [<span data-ttu-id="39fe3-135">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="39fe3-135">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="39fe3-136">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="39fe3-136">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="39fe3-137">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="39fe3-137">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="39fe3-138">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="39fe3-138">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="39fe3-139">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="39fe3-139">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

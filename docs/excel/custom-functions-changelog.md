---
ms.date: 01/08/2019
description: 发现 Excel 自定义函数的最新更新。
title: 自定义函数更改日志（预览）
ms.openlocfilehash: 48954ce759c7873925eb56a52d09b7196956542a
ms.sourcegitcommit: 9afcb1bb295ec0c8940ed3a8364dbac08ef6b382
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2019
ms.locfileid: "27773213"
---
# <a name="custom-functions-changelog-preview"></a><span data-ttu-id="2d001-103">自定义函数更改日志（预览）</span><span class="sxs-lookup"><span data-stu-id="2d001-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="2d001-104">Excel 自定义函数仍处于预览状态，这意味着将会对该产品进行频繁更改，包括更改和发布新功能。</span><span class="sxs-lookup"><span data-stu-id="2d001-104">Excel custom functions is still in preview and that means there are frequent changes to the product, including changes and the release of new features.</span></span> <span data-ttu-id="2d001-105">此更改日志提供了与产品所有更改相关的最新信息。</span><span class="sxs-lookup"><span data-stu-id="2d001-105">This changelog provides the most up-to-date information about any changes to the product.</span></span>

- <span data-ttu-id="2d001-106">**2017 年 11 月 7 日**：发布了\*自定义函数（预览）和示例</span><span class="sxs-lookup"><span data-stu-id="2d001-106">**Nov 7, 2017**: Shipped\* the custom functions preview and samples</span></span>
- <span data-ttu-id="2d001-107">**2017 年 11 月 20 日**：修复了使用内部版本 8801 及更高版本的函数的兼容性问题</span><span class="sxs-lookup"><span data-stu-id="2d001-107">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="2d001-108">**2017 年 11 月 28 日**：发布了\*对取消异步函数的支持（需要对流式处理函数进行相应更改）</span><span class="sxs-lookup"><span data-stu-id="2d001-108">**Nov 28, 2017**: Shipped\* support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="2d001-109">**2018 年 5 月 7 日**：发布了\*对 Mac、Excel Online 和在进程中运行的异步函数的支持</span><span class="sxs-lookup"><span data-stu-id="2d001-109">**May 7, 2018**: Shipped\* support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="2d001-110">**2018 年 9 月 20 日**：发布了对自定义函数 JavaScript 运行时的支持。</span><span class="sxs-lookup"><span data-stu-id="2d001-110">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="2d001-111">有关详细信息，请参阅 [Excel 自定义函数的运行时](custom-functions-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="2d001-111">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>
- <span data-ttu-id="2d001-112">**2018 年 10 月 20 日**：随着 [10 月预览体验内部版本](https://support.office.com/zh-CN/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24)的推出，自定义函数现在需要适用于 Windows Desktop 和 Online 的[自定义函数元数据](custom-functions-json.md)中的“id”参数。</span><span class="sxs-lookup"><span data-stu-id="2d001-112">**October 20, 2018**: With the [October Insiders build](https://support.office.com/zh-CN/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), Custom Functions now requires the 'id' parameter in your [custom functions metadata](custom-functions-json.md) for Windows Desktop and Online.</span></span> <span data-ttu-id="2d001-113">在 Mac 上，应忽略此参数。</span><span class="sxs-lookup"><span data-stu-id="2d001-113">On Mac, this parameter should be ignored.</span></span>
- <span data-ttu-id="2d001-114">**2018 年 12 月 12 日**：自定义函数中现在包括用于发现单元格地址的方法。</span><span class="sxs-lookup"><span data-stu-id="2d001-114">**December 12, 2018**: Custom functions now include a way to discover a cell's address.</span></span> <span data-ttu-id="2d001-115">有关详细信息，请参阅[确定调用自定义函数的单元格](custom-functions-overview.md#determine-which-cell-invoked-your-custom-function)。</span><span class="sxs-lookup"><span data-stu-id="2d001-115">For more information, see [Determine which cell invoked your custom function](custom-functions-overview.md#determine-which-cell-invoked-your-custom-function).</span></span>
- <span data-ttu-id="2d001-116">**2019 年 1 月 8 日**：绑定方法 `CustomFunctionMapping()` 已更改为 `CustomFunctions.associate()`。</span><span class="sxs-lookup"><span data-stu-id="2d001-116">**January 8, 2019**: Binding method `CustomFunctionMapping()` has been altered to `CustomFunctions.associate()`.</span></span> <span data-ttu-id="2d001-117">有关详细信息，请参阅[自定义函数最佳实践（预览）](custom-functions-best-practices.md)。</span><span class="sxs-lookup"><span data-stu-id="2d001-117">For more information, see [Custom functions best practices (preview)](custom-functions-best-practices.md).</span></span>

<span data-ttu-id="2d001-118">\* 转到 [Office 预览体验成员](https://products.office.com/office-insider)频道（以前称为“预览体验成员 - 快”）</span><span class="sxs-lookup"><span data-stu-id="2d001-118">\* to the [Office Insider](https://products.office.com/office-insider) channel (formerly called "Insider Fast")</span></span>

<span data-ttu-id="2d001-119">有关产品的已知问题列表，请参阅[已知问题](custom-functions-overview.md#known-issues)。</span><span class="sxs-lookup"><span data-stu-id="2d001-119">For a list of known issues with the product, see [Known Issues](custom-functions-overview.md#known-issues).</span></span> 

## <a name="see-also"></a><span data-ttu-id="2d001-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2d001-120">See also</span></span>

* [<span data-ttu-id="2d001-121">自定义函数概述</span><span class="sxs-lookup"><span data-stu-id="2d001-121">Custom functions overview</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="2d001-122">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="2d001-122">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="2d001-123">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="2d001-123">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="2d001-124">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="2d001-124">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="2d001-125">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="2d001-125">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

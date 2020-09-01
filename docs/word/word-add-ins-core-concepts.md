---
title: Word JavaScript API 基本编程概念
description: 使用 Word JavaScript API 生成适用于 Word 的加载项。
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: 1e7a90d4be378ed9b2c1f30ebebd4a0beec45a11
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293091"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a><span data-ttu-id="9857e-103">Word JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="9857e-103">Fundamental programming concepts with the Word JavaScript API</span></span>

<span data-ttu-id="9857e-104">本文介绍使用 [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) 生成适用于 Word 2016 或更高版本的加载项的基本概念。</span><span class="sxs-lookup"><span data-stu-id="9857e-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins for Word 2016 or later.</span></span>

## <a name="referencing-officejs"></a><span data-ttu-id="9857e-105">引用 Office.js</span><span class="sxs-lookup"><span data-stu-id="9857e-105">Referencing Office.js</span></span>

<span data-ttu-id="9857e-106">可以从以下位置引用 Office.js：</span><span class="sxs-lookup"><span data-stu-id="9857e-106">You can reference Office.js from the following locations:</span></span>

- <span data-ttu-id="9857e-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - 将此资源用于生产外接程序。</span><span class="sxs-lookup"><span data-stu-id="9857e-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - use this resource for production add-ins.</span></span>

- <span data-ttu-id="9857e-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - 通过此资源试用预览版功能。</span><span class="sxs-lookup"><span data-stu-id="9857e-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use this resource to try out preview features.</span></span>

## <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="9857e-109">Word JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="9857e-109">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="9857e-110">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="9857e-110">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="9857e-111">Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。</span><span class="sxs-lookup"><span data-stu-id="9857e-111">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="9857e-112">有关 Word JavaScript API 要求集的详细信息，请参阅 [Word JavaScript API 要求集](../reference/requirement-sets/word-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="9857e-112">For detailed information about Word JavaScript API requirement sets, see [Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md).</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="9857e-113">运行 Word 加载项</span><span class="sxs-lookup"><span data-stu-id="9857e-113">Running Word add-ins</span></span>

<span data-ttu-id="9857e-114">若要运行加载项，请使用 `Office.initialize` 事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="9857e-114">To run your add-in, use an `Office.initialize` event handler.</span></span> <span data-ttu-id="9857e-115">若要详细了解如何初始化加载项，请参阅[了解 API](../develop/understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="9857e-115">For more information about add-in initialization, see [Understanding the API](../develop/understanding-the-javascript-api-for-office.md).</span></span>

<span data-ttu-id="9857e-116">面向 Word 2016 或更高版本的加载项可以使用特定于 Word 的 API。</span><span class="sxs-lookup"><span data-stu-id="9857e-116">Add-ins that target Word 2016 or later can use the Word-specific APIs.</span></span> <span data-ttu-id="9857e-117">它们将 Word 交互逻辑作为函数传递到 `Word.run()` 方法中。</span><span class="sxs-lookup"><span data-stu-id="9857e-117">They pass the Word-interaction logic as a function into the `Word.run()` method.</span></span> <span data-ttu-id="9857e-118">请参阅[使用特定于应用程序的 API 模型](../develop/application-specific-api-model.md)，了解如何与此编程模型中的 Word 文档进行交互。</span><span class="sxs-lookup"><span data-stu-id="9857e-118">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about how to interact with the Word document in this programming model.</span></span>

<span data-ttu-id="9857e-119">以下示例显示如何使用 `Word.run()` 方法初始化和运行 Word 加载项。</span><span class="sxs-lookup"><span data-stu-id="9857e-119">The following example shows how to initialize and run a Word add-in by using the `Word.run()` method.</span></span>

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

## <a name="see-also"></a><span data-ttu-id="9857e-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9857e-120">See also</span></span>

- [<span data-ttu-id="9857e-121">Word JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="9857e-121">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="9857e-122">生成首个 Word 加载项</span><span class="sxs-lookup"><span data-stu-id="9857e-122">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="9857e-123">Word 加载项教程</span><span class="sxs-lookup"><span data-stu-id="9857e-123">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="9857e-124">Word JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="9857e-124">Word JavaScript API reference</span></span>](/javascript/api/word)

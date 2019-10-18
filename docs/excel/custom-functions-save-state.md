---
ms.date: 07/10/2019
description: 使用 `OfficeRuntime.storage` 保存自定义函数中的状态。
title: 保存并共享自定义函数中的状态
localization_priority: Priority
ms.openlocfilehash: a1b70433ef0c00d507175b32fc12603ff3de1e3f
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771587"
---
# <a name="save-and-share-state-in-custom-functions"></a><span data-ttu-id="ca791-103">保存并共享自定义函数中的状态</span><span class="sxs-lookup"><span data-stu-id="ca791-103">Save and share state in custom functions</span></span>

<span data-ttu-id="ca791-104">使用 `OfficeRuntime.storage` 对象保存与加载项中的自定义函数或任务窗格相关的状态。</span><span class="sxs-lookup"><span data-stu-id="ca791-104">Use the `OfficeRuntime.storage` object to save state related to custom functions or the task pane in your add-in.</span></span> <span data-ttu-id="ca791-105">存储限制为每个域 10 MB（可以在多个加载项中共享）。</span><span class="sxs-lookup"><span data-stu-id="ca791-105">Storage is limited to 10 MB per domain (which may be shared across multiple add-ins).</span></span> <span data-ttu-id="ca791-106">在 Windows 版 Excel 中，`storage` 对象是自定义函数运行时内的单独位置；但对于 Excel 网页版和 Mac 版 Excel，`storage` 对象与浏览器的 `localStorage` 相同。</span><span class="sxs-lookup"><span data-stu-id="ca791-106">In Excel on Windows, the `storage` object is a separate location within the custom functions runtime, but for Excel Online and Excel for Mac, the `storage` object is the same as the browser's `localStorage`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="ca791-107">可以通过多种方式使用 `storage` 进行状态管理：</span><span class="sxs-lookup"><span data-stu-id="ca791-107">There are multiple ways to use `storage` for state management:</span></span>

- <span data-ttu-id="ca791-108">可以存储自定义函数的默认值，以便在你离线和无法触及网页资源时使用。</span><span class="sxs-lookup"><span data-stu-id="ca791-108">You can store default values for custom functions to use when you are offline and unable to reach a web resource.</span></span>
- <span data-ttu-id="ca791-109">可以存储自定义函数值，以免额外调用网页资源。</span><span class="sxs-lookup"><span data-stu-id="ca791-109">You can save values for custom functions to use to avoid making additional calls to a web resource.</span></span>
- <span data-ttu-id="ca791-110">可以保存自定义函数中的值。</span><span class="sxs-lookup"><span data-stu-id="ca791-110">You can save values from your custom function.</span></span>
- <span data-ttu-id="ca791-111">可以存储任务窗格中的值。</span><span class="sxs-lookup"><span data-stu-id="ca791-111">You can store values from your task pane.</span></span>

<span data-ttu-id="ca791-112">以下代码示例演示了如何将项存储于 `storage` 中并检索它。</span><span class="sxs-lookup"><span data-stu-id="ca791-112">The following code sample illustrates how to store an item into `storage` and retrieve it.</span></span>

```js
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}
```

<span data-ttu-id="ca791-113">[GitHub 上的更详细代码示例](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)提供了将此信息传递到任务窗格的示例。</span><span class="sxs-lookup"><span data-stu-id="ca791-113">[A more detailed code sample on GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) gives an example of passing this information to the task pane.</span></span>

>[!NOTE]
> <span data-ttu-id="ca791-114">`storage` 对象取代了先前名为 `AsyncStorage` 的存储对象（现已启用）。</span><span class="sxs-lookup"><span data-stu-id="ca791-114">The `storage` object replaces the previous storage object named `AsyncStorage` which is now deprecated.</span></span> <span data-ttu-id="ca791-115">如果在当前的自定义函数代码中使用 `AsyncStorage` 对象，请将其更新为使用 `storage` 对象。</span><span class="sxs-lookup"><span data-stu-id="ca791-115">If using the `AsyncStorage` object in your current custom functions code, please update it to use the `storage` object.</span></span>

## <a name="next-steps"></a><span data-ttu-id="ca791-116">后续步骤</span><span class="sxs-lookup"><span data-stu-id="ca791-116">Next steps</span></span>
<span data-ttu-id="ca791-117">了解如何[为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="ca791-117">Learn how to [autogenerate the JSON metadata for your custom functions](custom-functions-json-autogeneration.md).</span></span> 

## <a name="see-also"></a><span data-ttu-id="ca791-118">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ca791-118">See also</span></span>

* [<span data-ttu-id="ca791-119">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="ca791-119">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="ca791-120">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="ca791-120">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="ca791-121">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="ca791-121">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="ca791-122">自定义函数调试</span><span class="sxs-lookup"><span data-stu-id="ca791-122">Custom functions debugging</span></span>](custom-functions-debugging.md)

---
ms.date: 05/17/2020
description: 了解不使用任务窗格及其特定 JavaScript 运行时的 Excel 自定义函数。
title: 不带 UI 的 Excel 自定义函数的运行时
localization_priority: Normal
ms.openlocfilehash: 31044d4569d230e252c05a39785fc7d47b802e37
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278355"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a><span data-ttu-id="a8b6b-103">不带 UI 的 Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="a8b6b-103">Runtime for UI-less Excel custom functions</span></span>

<span data-ttu-id="a8b6b-104">不使用任务窗格的自定义函数（不带 UI 的自定义函数）使用旨在优化计算性能的 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-104">Custom functions that don't use a task pane (UI-less custom functions) use a JavaScript runtime that is designed to optimize performance of calculations.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="a8b6b-105">此 JavaScript 运行时提供对命名空间中的 Api 的访问 `OfficeRuntime` ，这些 api 可由无 UI 的自定义函数和任务窗格用于存储数据。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-105">This JavaScript runtime provides access to APIs in the `OfficeRuntime` namespace that can be used by UI-less custom functions and the task pane to store data.</span></span>

## <a name="requesting-external-data"></a><span data-ttu-id="a8b6b-106">请求外部数据</span><span class="sxs-lookup"><span data-stu-id="a8b6b-106">Requesting external data</span></span>

<span data-ttu-id="a8b6b-107">在无 UI 的自定义函数中，您可以通过使用 API （如[Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ）或使用[XmlHttpRequest （XHR）](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)来请求外部数据，这是一个标准 web API，它会发出 HTTP 请求，以与服务器交互。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-107">Within a UI-less custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="a8b6b-108">请注意，在进行 XmlHttpRequests 时，无 UI 的函数必须使用其他安全措施，这需要[相同的源策略](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)和简单的[CORS](https://www.w3.org/TR/cors/)。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-108">Be aware that UI-less functions must use additional security measures when making XmlHttpRequests, requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="a8b6b-109">简单的 CORS 实现不能使用 cookie，仅支持简单方法（GET、HEAD、POST）。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-109">A simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="a8b6b-110">简单的 CORS 接受字段名称为 `Accept`、`Accept-Language`、`Content-Language` 的简单标题。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-110">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="a8b6b-111">您还可以使用 `Content-Type` 简单 CORS 中的标头，只要内容类型为 `application/x-www-form-urlencoded` 、 `text/plain` 或 `multipart/form-data` 。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-111">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

## <a name="storing-and-accessing-data"></a><span data-ttu-id="a8b6b-112">存储和访问数据</span><span class="sxs-lookup"><span data-stu-id="a8b6b-112">Storing and accessing data</span></span>

<span data-ttu-id="a8b6b-113">在不带 UI 的自定义函数中，您可以使用对象存储和访问数据 `OfficeRuntime.storage` 。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-113">Within a UI-less custom function, you can store and access data by using the `OfficeRuntime.storage` object.</span></span> <span data-ttu-id="a8b6b-114">`Storage`是一个永久性的未加密的键值存储系统，可提供[localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage)的替代方法，不能由无 UI 的自定义函数使用。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-114">`Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used by UI-less custom functions.</span></span> <span data-ttu-id="a8b6b-115">`Storage`每个域提供 10 MB 的数据。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-115">`Storage` offers 10 MB of data per domain.</span></span> <span data-ttu-id="a8b6b-116">域可由多个加载项共享。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-116">Domains can be shared by more than one add-in.</span></span>

<span data-ttu-id="a8b6b-117">`Storage` 旨在作为共享存储解决方案，这意味着外接程序的多个部分将能访问相同数据。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-117">`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="a8b6b-118">例如，可以将用于用户身份验证的令牌存储在中， `storage` 因为无 UI 的自定义函数和外接程序 ui 元素（如任务窗格）可以访问它。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-118">For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a UI-less custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="a8b6b-119">同样，如果两个加载项共享同一个域（例如， `www.contoso.com/addin1` ， `www.contoso.com/addin2` ），则也允许它们前后共享信息 `storage` 。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-119">Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through `storage`.</span></span> <span data-ttu-id="a8b6b-120">请注意，具有不同子域的外接程序将具有不同的实例 `storage` （例如， `subdomain.contoso.com/addin1` `differentsubdomain.contoso.com/addin2` ）。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-120">Note that add-ins which have different subdomains will have different instances of `storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).</span></span>

<span data-ttu-id="a8b6b-121">由于 `storage` 可能是共享的位置，因此一定要认识到，可能会存在替代键值对的情况。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-121">Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="a8b6b-122">`storage` 对象支持以下方法：</span><span class="sxs-lookup"><span data-stu-id="a8b6b-122">The following methods are available on the `storage` object:</span></span>

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

<span data-ttu-id="a8b6b-123">.</span><span class="sxs-lookup"><span data-stu-id="a8b6b-123">.</span></span>[!NOTE]
> <span data-ttu-id="a8b6b-124">没有用于清除所有信息的方法（例如 `clear` ）。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-124">There's no method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="a8b6b-125">相反，需要使用 `removeItems` 来一次性删除多个条目。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-125">Instead, you should instead use `removeItems` to remove multiple entries at a time.</span></span>

### <a name="officeruntimestorage-example"></a><span data-ttu-id="a8b6b-126">OfficeRuntime 示例</span><span class="sxs-lookup"><span data-stu-id="a8b6b-126">OfficeRuntime.storage example</span></span>

<span data-ttu-id="a8b6b-127">下面的代码示例调用 `OfficeRuntime.storage.setItem` 函数，以将键和值设置为 `storage` 。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-127">The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.</span></span>

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="a8b6b-128">其他注意事项</span><span class="sxs-lookup"><span data-stu-id="a8b6b-128">Additional considerations</span></span>

<span data-ttu-id="a8b6b-129">如果外接程序仅使用无 UI 的自定义函数，请注意，不能使用不带 UI 的自定义函数访问文档对象模型（DOM），也不能使用依赖 DOM 的 jQuery 等库。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-129">If your add-in only uses UI-less custom functions, note that you can't access the Document Object Model (DOM) with UI-less custom functions or use libraries like jQuery that rely on the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a8b6b-130">后续步骤</span><span class="sxs-lookup"><span data-stu-id="a8b6b-130">Next steps</span></span>
<span data-ttu-id="a8b6b-131">了解如何[调试不带 UI 的自定义函数](custom-functions-debugging.md)。</span><span class="sxs-lookup"><span data-stu-id="a8b6b-131">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a8b6b-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a8b6b-132">See also</span></span>

* [<span data-ttu-id="a8b6b-133">对 UI 进行身份验证-更少的自定义函数</span><span class="sxs-lookup"><span data-stu-id="a8b6b-133">Authenticate UI-less custom functions</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="a8b6b-134">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="a8b6b-134">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="a8b6b-135">自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="a8b6b-135">Custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

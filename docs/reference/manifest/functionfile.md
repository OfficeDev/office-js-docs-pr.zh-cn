---
title: 清单文件中的 FunctionFile 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eec1dc8eb2e099670469af6ef300592fc4a31e64
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324867"
---
# <a name="functionfile-element"></a><span data-ttu-id="c9668-102">FunctionFile 元素</span><span class="sxs-lookup"><span data-stu-id="c9668-102">FunctionFile element</span></span>

<span data-ttu-id="c9668-103">为外接程序通过外接程序命令公开的操作指定源代码文件，这些外接程序命令执行 JavaScript 函数，而不显示 UI。</span><span class="sxs-lookup"><span data-stu-id="c9668-103">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI.</span></span> <span data-ttu-id="c9668-104">`FunctionFile`元素是[DesktopFormFactor](desktopformfactor.md)或[MobileFormFactor](mobileformfactor.md)的子元素。</span><span class="sxs-lookup"><span data-stu-id="c9668-104">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="c9668-105">`FunctionFile`元素`resid`的属性设置为`id` `Url` `Resources`元素中元素的属性的值，该元素包含的 HTML 文件的 URL 包含或加载由[Control 元素](control.md)定义的无用户界面外接程序命令按钮所使用的所有 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="c9668-105">The `resid` attribute of the `FunctionFile` element is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="c9668-106">以下是`FunctionFile`元素的示例。</span><span class="sxs-lookup"><span data-stu-id="c9668-106">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="c9668-107">`FunctionFile`元素指示的 HTML 文件中的 JavaScript 必须调用`Office.initialize`和定义采用单个参数的命名函数： `event`。</span><span class="sxs-lookup"><span data-stu-id="c9668-107">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="c9668-108">这些函数应使用 `item.notificationMessages` API 向用户指示进度、成功或失败。</span><span class="sxs-lookup"><span data-stu-id="c9668-108">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="c9668-109">此外，它应在完成执行时调用 `event.completed`。</span><span class="sxs-lookup"><span data-stu-id="c9668-109">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="c9668-110">函数的名称在无 UI 按钮的`FunctionName`元素中使用。</span><span class="sxs-lookup"><span data-stu-id="c9668-110">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="c9668-111">下面是定义`trackMessage`函数的 HTML 文件的一个示例。</span><span class="sxs-lookup"><span data-stu-id="c9668-111">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

<span data-ttu-id="c9668-112">下面的代码演示如何实现由`FunctionName`使用的函数。</span><span class="sxs-lookup"><span data-stu-id="c9668-112">The following code shows how to implement the function used by `FunctionName`.</span></span>

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> <span data-ttu-id="c9668-113">对事件进行`event.completed`信号的调用，以表示已成功处理事件。</span><span class="sxs-lookup"><span data-stu-id="c9668-113">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="c9668-114">当某个函数被多次调用时（例如在同一外接程序命令上进行多次单击），所有事件将自动排队。</span><span class="sxs-lookup"><span data-stu-id="c9668-114">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="c9668-115">第一个事件将自动运行，而其他事件仍保持在队列中。</span><span class="sxs-lookup"><span data-stu-id="c9668-115">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="c9668-116">当函数调用`event.completed`时，将运行对该函数的下一个排队调用。</span><span class="sxs-lookup"><span data-stu-id="c9668-116">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="c9668-117">您必须调用`event.completed`;否则，您的函数将不会运行。</span><span class="sxs-lookup"><span data-stu-id="c9668-117">You must call `event.completed`; otherwise your function will not run.</span></span>

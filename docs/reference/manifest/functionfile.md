---
title: 清单文件中的 FunctionFile 元素
description: 为外接程序通过外接程序命令公开的操作指定源代码文件，这些外接程序命令执行 JavaScript 函数，而不显示 UI。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 44bfd514025b8a23f4f6acdf3fec004485ca4c5a
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771392"
---
# <a name="functionfile-element"></a><span data-ttu-id="08010-103">FunctionFile 元素</span><span class="sxs-lookup"><span data-stu-id="08010-103">FunctionFile element</span></span>

<span data-ttu-id="08010-104">指定外接程序通过以下方式之一公开的操作的源代码文件：</span><span class="sxs-lookup"><span data-stu-id="08010-104">Specifies the source code file for operations that an add-in exposes in one of the following ways:</span></span>

* <span data-ttu-id="08010-105">执行 JavaScript 函数而不是显示 UI 的外接程序命令。</span><span class="sxs-lookup"><span data-stu-id="08010-105">Add-in commands that execute a JavaScript function instead of displaying UI.</span></span>
* <span data-ttu-id="08010-106">执行 JavaScript 函数的键盘快捷方式。</span><span class="sxs-lookup"><span data-stu-id="08010-106">Keyboard shortcuts that execute a JavaScript function.</span></span>

<span data-ttu-id="08010-107">该 `FunctionFile` 元素是 [DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor 的子元素](mobileformfactor.md)。</span><span class="sxs-lookup"><span data-stu-id="08010-107">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="08010-108">元素的属性不能超过 32 个字符，并设置为元素中某个元素的属性值，该元素包含 HTML 文件的 URL，该文件包含或加载无 UI 加载项命令按钮使用的所有 `resid` `FunctionFile` `id` `Url` `Resources` JavaScript[](control.md)函数，如 Control 元素所定义。</span><span class="sxs-lookup"><span data-stu-id="08010-108">The `resid` attribute of the `FunctionFile` element can be no more than 32 characters and is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="08010-109">下面是该元素 `FunctionFile` 的一个示例。</span><span class="sxs-lookup"><span data-stu-id="08010-109">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="08010-110">元素指示的 HTML 文件中 JavaScript 必须调用并定义使用单个参数的命名 `FunctionFile` `Office.initialize` 函数 `event` ：</span><span class="sxs-lookup"><span data-stu-id="08010-110">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="08010-111">这些函数应使用 `item.notificationMessages` API 向用户指示进度、成功或失败。</span><span class="sxs-lookup"><span data-stu-id="08010-111">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="08010-112">此外，它应在完成执行时调用 `event.completed`。</span><span class="sxs-lookup"><span data-stu-id="08010-112">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="08010-113">这些函数的名称在无 `FunctionName` UI 按钮的元素中使用。</span><span class="sxs-lookup"><span data-stu-id="08010-113">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="08010-114">下面是定义函数的 HTML 文件 `trackMessage` 的示例。</span><span class="sxs-lookup"><span data-stu-id="08010-114">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

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

<span data-ttu-id="08010-115">下面的代码演示如何实现所使用的函数 `FunctionName` 。</span><span class="sxs-lookup"><span data-stu-id="08010-115">The following code shows how to implement the function used by `FunctionName`.</span></span>

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
> <span data-ttu-id="08010-116">对 `event.completed` 信号的调用，指示已成功处理事件。</span><span class="sxs-lookup"><span data-stu-id="08010-116">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="08010-117">当某个函数被多次调用时（例如在同一外接程序命令上进行多次单击），所有事件将自动排队。</span><span class="sxs-lookup"><span data-stu-id="08010-117">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="08010-118">第一个事件将自动运行，而其他事件仍保持在队列中。</span><span class="sxs-lookup"><span data-stu-id="08010-118">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="08010-119">函数调用时 `event.completed` ，对函数的下一个排队调用将运行。</span><span class="sxs-lookup"><span data-stu-id="08010-119">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="08010-120">您必须调用 `event.completed` ;否则函数将不会运行。</span><span class="sxs-lookup"><span data-stu-id="08010-120">You must call `event.completed`; otherwise your function will not run.</span></span>

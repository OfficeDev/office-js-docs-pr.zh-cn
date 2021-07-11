---
title: 清单文件中的 FunctionFile 元素
description: 为外接程序通过外接程序命令公开的操作指定源代码文件，这些外接程序命令执行 JavaScript 函数，而不显示 UI。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: f31a1bc7a561305a89f5388102a4985aaa31fe37
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348298"
---
# <a name="functionfile-element"></a>FunctionFile 元素

指定外接程序通过以下方法之一公开的操作的源代码文件。

* 执行 JavaScript 函数而不是显示 UI 的外接程序命令。
* 执行 JavaScript 函数的键盘快捷方式。

元素 `FunctionFile` 是 [DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor 的子元素](mobileformfactor.md)。 元素的 属性不能超过 32 个字符，并且设置为 元素中元素的 属性值，该元素包含 HTML 文件的 URL，该文件包含或加载无 UI 加载项命令按钮使用的所有 `resid` `FunctionFile` `id` `Url` `Resources` JavaScript[](control.md)函数，如 Control 元素所定义。

下面是 元素 `FunctionFile` 的一个示例。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

元素指示的 HTML 文件中 JavaScript 必须调用并定义使用单个参数 `FunctionFile` `Office.initialize` 的命名函数 `event` ：。 这些函数应使用 `item.notificationMessages` API 向用户指示进度、成功或失败。 此外，它应在完成执行时调用 `event.completed`。 这些函数的名称在 无 UI `FunctionName` 按钮的 元素中使用。

下面是定义函数的 HTML 文件 `trackMessage` 的示例。

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

以下代码演示如何实现 由 使用的函数 `FunctionName` 。

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
> 对信号 `event.completed` 的调用表示已成功处理事件。 当某个函数被多次调用时（例如在同一外接程序命令上进行多次单击），所有事件将自动排队。 第一个事件将自动运行，而其他事件仍保持在队列中。 函数调用 `event.completed` 时，将运行对函数的下一个排队调用。 您必须调用 `event.completed` ;否则函数将不会运行。

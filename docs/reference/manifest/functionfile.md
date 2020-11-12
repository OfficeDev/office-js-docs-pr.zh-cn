---
title: 清单文件中的 FunctionFile 元素
description: 为外接程序通过外接程序命令公开的操作指定源代码文件，这些外接程序命令执行 JavaScript 函数，而不显示 UI。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 4c47c3e4b824f2b93aaea17cef88e01f748d6f95
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996443"
---
# <a name="functionfile-element"></a>FunctionFile 元素

按以下方式之一，指定外接程序所公开的操作的源代码文件：

* 执行 JavaScript 函数而不是显示 UI 的外接程序命令。
* 执行 JavaScript 函数的键盘快捷方式。

`FunctionFile`元素是[DesktopFormFactor](desktopformfactor.md)或[MobileFormFactor](mobileformfactor.md)的子元素。 `resid`元素的属性 `FunctionFile` 设置为元素中元素的属性的值，该元素包含的 `id` `Url` `Resources` HTML 文件的 URL 包含或加载由[Control 元素](control.md)定义的无用户界面外接程序命令按钮所使用的所有 JavaScript 函数。

以下是元素的示例 `FunctionFile` 。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

元素指示的 HTML 文件中的 JavaScript `FunctionFile` 必须调用 `Office.initialize` 和定义采用单个参数的命名函数： `event` 。 这些函数应使用 `item.notificationMessages` API 向用户指示进度、成功或失败。 此外，它应在完成执行时调用 `event.completed`。 函数的名称在无 `FunctionName` UI 按钮的元素中使用。

下面是定义函数的 HTML 文件的一个示例 `trackMessage` 。

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

下面的代码演示如何实现由使用的函数 `FunctionName` 。

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
> 对事件进行 `event.completed` 信号的调用，以表示已成功处理事件。 当某个函数被多次调用时（例如在同一外接程序命令上进行多次单击），所有事件将自动排队。 第一个事件将自动运行，而其他事件仍保持在队列中。 当函数调用时 `event.completed` ，将运行对该函数的下一个排队调用。 您必须调用 `event.completed` ; 否则您的函数将不会运行。

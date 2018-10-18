# <a name="functionfile-element"></a>FunctionFile 元素

为外接程序通过外接程序命令公开的操作指定源代码文件，这些外接程序命令执行 JavaScript 函数，而不显示 UI。**FunctionFile** 元素是 [DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor](mobileformfactor.md) 的子元素。**FunctionFile** 元素的 **resid** 属性被设置为 **Resources** 元素中的 **Url** 元素的 **id** 属性值，Resources 元素包含 HTML 文件的 URL，其中包含或加载所有由无 UI 外接程序命令按钮使用的 JavaScript 函数（由 [Control](control.md) 元素定义）。

下面的示例展示了 **FunctionFile** 元素。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

指示由 **FunctionFile** 元素的 HTML 文件中的 JavaScript 必须调用 `Office.initialize` ，并定义单个参数命名的函数： `event`。 应使用函数 `item.notificationMessages` API 显示进度、 成功或对用户失败。 它还应调用 `event.completed` 执行已完成。 无用户界面按钮， **FunctionName** 元素中使用的函数名称。

以下是定义 **trackMessage** 函数的 HTML 文件的示例。

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

下面的代码展示了如何实现 **FunctionName** 使用的函数。

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
> 重要说明  调用 **event.completed** 表示你已成功处理了事件。 当某个函数被多次调用时（例如多次单击同一加载项命令时），所有事件将自动排队。 首个事件将自动运行，而其他事件仍保持在队列中。 当函数调用 **event.completed** 时，将运行队列中下一个对此函数的调用。 必须实现 **event.completed**，否则函数将不会运行。
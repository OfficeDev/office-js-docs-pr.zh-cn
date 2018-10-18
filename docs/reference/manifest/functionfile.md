# <a name="functionfile-element"></a><span data-ttu-id="454ed-101">FunctionFile 元素</span><span class="sxs-lookup"><span data-stu-id="454ed-101">FunctionFile element</span></span>

<span data-ttu-id="454ed-p101">为外接程序通过外接程序命令公开的操作指定源代码文件，这些外接程序命令执行 JavaScript 函数，而不显示 UI。**FunctionFile** 元素是 [DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor](mobileformfactor.md) 的子元素。**FunctionFile** 元素的 **resid** 属性被设置为 **Resources** 元素中的 **Url** 元素的 **id** 属性值，Resources 元素包含 HTML 文件的 URL，其中包含或加载所有由无 UI 外接程序命令按钮使用的 JavaScript 函数（由 [Control](control.md) 元素定义）。</span><span class="sxs-lookup"><span data-stu-id="454ed-p101">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="454ed-105">下面的示例展示了 **FunctionFile** 元素。</span><span class="sxs-lookup"><span data-stu-id="454ed-105">The following is an example of the **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="454ed-106">指示由 **FunctionFile** 元素的 HTML 文件中的 JavaScript 必须调用 `Office.initialize` ，并定义单个参数命名的函数： `event`。</span><span class="sxs-lookup"><span data-stu-id="454ed-106">The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="454ed-107">应使用函数 `item.notificationMessages` API 显示进度、 成功或对用户失败。</span><span class="sxs-lookup"><span data-stu-id="454ed-107">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="454ed-108">它还应调用 `event.completed` 执行已完成。</span><span class="sxs-lookup"><span data-stu-id="454ed-108">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="454ed-109">无用户界面按钮， **FunctionName** 元素中使用的函数名称。</span><span class="sxs-lookup"><span data-stu-id="454ed-109">The name of the functions are used in the **FunctionName** element for UI-less buttons.</span></span>

<span data-ttu-id="454ed-110">以下是定义 **trackMessage** 函数的 HTML 文件的示例。</span><span class="sxs-lookup"><span data-stu-id="454ed-110">The following is an example of an HTML file defining a **trackMessage** function.</span></span>

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

<span data-ttu-id="454ed-111">下面的代码展示了如何实现 **FunctionName** 使用的函数。</span><span class="sxs-lookup"><span data-stu-id="454ed-111">The following code shows how to implement the function used by **FunctionName**.</span></span>

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
> <span data-ttu-id="454ed-112">重要说明  调用 **event.completed** 表示你已成功处理了事件。</span><span class="sxs-lookup"><span data-stu-id="454ed-112">IMPORTANT  The call to **event.completed** signals that you have successfully handled the event.</span></span> <span data-ttu-id="454ed-113">当某个函数被多次调用时（例如多次单击同一加载项命令时），所有事件将自动排队。</span><span class="sxs-lookup"><span data-stu-id="454ed-113">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="454ed-114">首个事件将自动运行，而其他事件仍保持在队列中。</span><span class="sxs-lookup"><span data-stu-id="454ed-114">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="454ed-115">当函数调用 **event.completed** 时，将运行队列中下一个对此函数的调用。</span><span class="sxs-lookup"><span data-stu-id="454ed-115">When your function calls **event.completed**, the next queued call to that function runs.</span></span> <span data-ttu-id="454ed-116">必须实现 **event.completed**，否则函数将不会运行。</span><span class="sxs-lookup"><span data-stu-id="454ed-116">You must implement **event.completed**, otherwise your function will not run.</span></span>
---
title: 在共享运行时中显示或隐藏 Office 加载项
description: 了解如何在连续运行时以编程方式隐藏或显示外接程序的用户界面
ms.date: 05/17/2020
localization_priority: Normal
ms.openlocfilehash: e49c47c86a986c85ad12e09666b7ac2fb5411322
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275712"
---
# <a name="show-or-hide-an-office-add-in-in-a-shared-runtime"></a><span data-ttu-id="81ab6-103">在共享运行时中显示或隐藏 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="81ab6-103">Show or hide an Office Add-in in a shared runtime</span></span>

<span data-ttu-id="81ab6-104">Office 外接程序可以包含以下任何部分：</span><span class="sxs-lookup"><span data-stu-id="81ab6-104">An Office Add-in can include any of the following parts:</span></span>

- <span data-ttu-id="81ab6-105">任务窗格</span><span class="sxs-lookup"><span data-stu-id="81ab6-105">A task pane</span></span>
- <span data-ttu-id="81ab6-106">无 UI 的函数文件（不使用任务窗格或其他用户界面元素的自定义函数）</span><span class="sxs-lookup"><span data-stu-id="81ab6-106">A UI-less function file (custom functions which do not use a task pane or other user interface elements)</span></span>
- <span data-ttu-id="81ab6-107">Excel 自定义函数</span><span class="sxs-lookup"><span data-stu-id="81ab6-107">An Excel custom function</span></span>

<span data-ttu-id="81ab6-108">默认情况下，每个部件都在自己的独立 JavaScript 运行时中运行，其中包含其自己的全局对象和全局变量。</span><span class="sxs-lookup"><span data-stu-id="81ab6-108">By default, each part runs in its own separate JavaScript runtime, with its own global object and global variables.</span></span>

<span data-ttu-id="81ab6-109">具有两个或更多个部件的外接程序可以共享一个通用的 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="81ab6-109">It's possible for add-ins with two or more parts to share a common JavaScript runtime.</span></span> <span data-ttu-id="81ab6-110">此共享运行时功能启用在外接程序运行时隐藏和重新打开任务窗格的新 Api。</span><span class="sxs-lookup"><span data-stu-id="81ab6-110">This shared runtime feature enables new APIs that hide and reopen the task pane while the add-in runs.</span></span>

## <a name="configure-an-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="81ab6-111">将外接程序配置为使用共享运行时</span><span class="sxs-lookup"><span data-stu-id="81ab6-111">Configure an add-in to use a shared runtime</span></span>

<span data-ttu-id="81ab6-112">若要将外接程序配置为使用共享运行时，请参阅[configure The Office 外接程序以使用共享运行时](configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="81ab6-112">To configure the add-in to use a shared runtime, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="show-and-hide-the-task-pane"></a><span data-ttu-id="81ab6-113">显示和隐藏任务窗格</span><span class="sxs-lookup"><span data-stu-id="81ab6-113">Show and hide the task pane</span></span>

<span data-ttu-id="81ab6-114">新的 Api 位于属性中 `Office.addin` 。</span><span class="sxs-lookup"><span data-stu-id="81ab6-114">The new APIs are in the `Office.addin` property.</span></span> <span data-ttu-id="81ab6-115">若要显示任务窗格，您的代码将调用 `Office.addin.showAsTaskpane()` 。</span><span class="sxs-lookup"><span data-stu-id="81ab6-115">To show the task pane, your code calls `Office.addin.showAsTaskpane()`.</span></span> <span data-ttu-id="81ab6-116">Office 将在任务窗格中显示分配给任务窗格的资源 ID （）的页面 `resid` 。</span><span class="sxs-lookup"><span data-stu-id="81ab6-116">Office will display in a task pane the page that you assigned to the resource ID (`resid`) for the task pane.</span></span> <span data-ttu-id="81ab6-117">这是 `resid` 分配给 `<SourceLocation>` 清单中的的的 `<Action xsi:type="ShowTaskpane">` 。</span><span class="sxs-lookup"><span data-stu-id="81ab6-117">This is the `resid` that you assigned to the `<SourceLocation>` of the `<Action xsi:type="ShowTaskpane">` in the manifest.</span></span> <span data-ttu-id="81ab6-118">（请参阅[配置 Office 外接程序以使用共享运行时](configure-your-add-in-to-use-a-shared-runtime.md)。）</span><span class="sxs-lookup"><span data-stu-id="81ab6-118">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).)</span></span>

<span data-ttu-id="81ab6-119">这是一种异步方法，因此，如果后续代码在完成之前不应运行，则代码应等待它。</span><span class="sxs-lookup"><span data-stu-id="81ab6-119">This is an asynchronous method, so your code should await it when the subsequent code should not run until it completes.</span></span> <span data-ttu-id="81ab6-120">使用关键字或方法等待这一完成 `await` `then()` ，具体取决于您使用的 JavaScript 语法。</span><span class="sxs-lookup"><span data-stu-id="81ab6-120">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span> <span data-ttu-id="81ab6-121">以下示例假定有一个名为**CurrentQuarterSales**的 Excel 工作表。</span><span class="sxs-lookup"><span data-stu-id="81ab6-121">The following assumes that there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="81ab6-122">每当激活此工作表时，加载项都应显示任务窗格。</span><span class="sxs-lookup"><span data-stu-id="81ab6-122">The add-in should make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="81ab6-123">该方法 `onCurrentQuarter` 是已为工作表注册的[onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated)事件的处理程序。</span><span class="sxs-lookup"><span data-stu-id="81ab6-123">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) event which has been registered for the worksheet.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="81ab6-124">若要隐藏任务窗格，您的代码将调用 `Office.addin.hide()` 。</span><span class="sxs-lookup"><span data-stu-id="81ab6-124">To hide the task pane, your code calls `Office.addin.hide()`.</span></span> <span data-ttu-id="81ab6-125">下面的示例是一个为[onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated)事件注册的处理程序。</span><span class="sxs-lookup"><span data-stu-id="81ab6-125">The following example is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated) event.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="81ab6-126">保留状态和事件侦听器</span><span class="sxs-lookup"><span data-stu-id="81ab6-126">Preservation of state and event listeners</span></span>

<span data-ttu-id="81ab6-127">`hide()`和 `showAsTaskpane()` 方法仅更改任务窗格的*可见性*。</span><span class="sxs-lookup"><span data-stu-id="81ab6-127">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="81ab6-128">它们不会卸载或重新加载它（或重新初始化其状态）。</span><span class="sxs-lookup"><span data-stu-id="81ab6-128">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="81ab6-129">请考虑以下方案：任务窗格是用选项卡设计的。</span><span class="sxs-lookup"><span data-stu-id="81ab6-129">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="81ab6-130">首次启动加载项时，"**主页**" 选项卡处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="81ab6-130">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="81ab6-131">假设用户打开 "**设置**" 选项卡，随后，任务窗格中的代码将调用 `hide()` 以响应某个事件。</span><span class="sxs-lookup"><span data-stu-id="81ab6-131">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="81ab6-132">仍在以后的代码调用， `showAsTaskpane()` 以响应另一个事件。</span><span class="sxs-lookup"><span data-stu-id="81ab6-132">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="81ab6-133">任务窗格将重新显示，并且 "**设置**" 选项卡仍处于选中状态。</span><span class="sxs-lookup"><span data-stu-id="81ab6-133">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![任务窗格的屏幕截图，其中有四个标签为 "主页"、"设置"、"收藏夹" 和 "帐户"。](../images/TaskpaneWithTabs.png)

<span data-ttu-id="81ab6-135">此外，即使任务窗格处于隐藏状态，在任务窗格中注册的任何事件侦听器也将继续运行。</span><span class="sxs-lookup"><span data-stu-id="81ab6-135">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="81ab6-136">请考虑以下方案：任务窗格有一个 Excel `Worksheet.onActivated` 和一个 `Worksheet.onDeactivated` 名为**Sheet1**的工作表的事件的已注册处理程序。</span><span class="sxs-lookup"><span data-stu-id="81ab6-136">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="81ab6-137">激活的处理程序导致在任务窗格中显示一个绿色点。</span><span class="sxs-lookup"><span data-stu-id="81ab6-137">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="81ab6-138">已停用的处理程序会将点变为红色（这是其默认状态）。</span><span class="sxs-lookup"><span data-stu-id="81ab6-138">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="81ab6-139">假设该代码 `hide()` 在**Sheet1**未激活且点为红色时调用。</span><span class="sxs-lookup"><span data-stu-id="81ab6-139">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="81ab6-140">在任务窗格处于隐藏状态时， **Sheet1**处于激活状态。</span><span class="sxs-lookup"><span data-stu-id="81ab6-140">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="81ab6-141">后续代码调用 `showAsTaskpane()` 以响应某个事件。</span><span class="sxs-lookup"><span data-stu-id="81ab6-141">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="81ab6-142">任务窗格打开时，点为绿色，因为即使任务窗格被隐藏，也会运行事件侦听器和处理程序。</span><span class="sxs-lookup"><span data-stu-id="81ab6-142">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

### <a name="handle-visibility-changed-event"></a><span data-ttu-id="81ab6-143">处理可见性更改事件</span><span class="sxs-lookup"><span data-stu-id="81ab6-143">Handle visibility changed event</span></span>

<span data-ttu-id="81ab6-144">当您的代码通过 or 更改任务窗格的可见性时 `showAsTaskpane()` `hide()` ，Office 将触发该 `VisibilityModeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="81ab6-144">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="81ab6-145">处理此事件可能很有用。</span><span class="sxs-lookup"><span data-stu-id="81ab6-145">It can be useful to handle this event.</span></span> <span data-ttu-id="81ab6-146">例如，假设任务窗格显示工作簿中所有工作表的列表。</span><span class="sxs-lookup"><span data-stu-id="81ab6-146">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="81ab6-147">如果在任务窗格处于隐藏状态时添加了一个新的工作表，使任务窗格可见，则它本身不会将新的工作表名称添加到列表中。</span><span class="sxs-lookup"><span data-stu-id="81ab6-147">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="81ab6-148">但您的代码可以响应 `VisibilityModeChanged` 事件以重新加载工作簿中所有工作表的[Worksheet.name](/javascript/api/excel/excel.worksheet#name)属性[。工作表](/javascript/api/excel/excel.workbook#worksheets)集合，如下面的示例代码所示。</span><span class="sxs-lookup"><span data-stu-id="81ab6-148">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="81ab6-149">若要注册事件的处理程序，请不要像在大多数 Office JavaScript 上下文中那样使用 "添加处理程序" 方法。</span><span class="sxs-lookup"><span data-stu-id="81ab6-149">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="81ab6-150">相反，有一个特殊的函数，您可以将其传递给处理程序： [onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-)。</span><span class="sxs-lookup"><span data-stu-id="81ab6-150">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="81ab6-151">示例如下。</span><span class="sxs-lookup"><span data-stu-id="81ab6-151">The following is an example.</span></span> <span data-ttu-id="81ab6-152">请注意， `args.visibilityMode` 属性的类型为[VisibilityMode](/javascript/api/office/office.visibilitymode)。</span><span class="sxs-lookup"><span data-stu-id="81ab6-152">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="81ab6-153">函数返回*deregisters*处理程序的另一个函数。</span><span class="sxs-lookup"><span data-stu-id="81ab6-153">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="81ab6-154">下面是一个简单但不可靠的示例：</span><span class="sxs-lookup"><span data-stu-id="81ab6-154">Here is a simple, but not robust, example:</span></span>

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="81ab6-155">`onVisibilityModeChanged`方法是异步的，这意味着，如果代码调用返回的取消*注册*处理程序 `onVisibilityModeChanged` ，则应确保在 `onVisibilityModeChanged` 调用取消注册处理程序之前已完成。</span><span class="sxs-lookup"><span data-stu-id="81ab6-155">The `onVisibilityModeChanged` method is asynchronous which means that if your code calls the *deregister* handler that `onVisibilityModeChanged` returns, you should ensure that `onVisibilityModeChanged` has completed before calling the deregister handler.</span></span> <span data-ttu-id="81ab6-156">执行此操作的一种方法是在 `await` 方法调用中使用关键字，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="81ab6-156">One way to do that is to use the `await` keyword on the method call as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="81ab6-157">如果您只想使用 ES2015 JavaScript，则代码可以使用 `then` 方法等待，直到返回的承诺对象已解决，并将返回的函数分配给全局变量，如以下示例中所示。</span><span class="sxs-lookup"><span data-stu-id="81ab6-157">If you want to use only pre-ES2015 JavaScript, your code can use the `then` method to wait until the returned Promise object has resolved and assign the returned function to a global variable as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler;

Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
}).then(function(removeHandler) {
        removeVisibilityModeHandler = removeHandler;
    });

// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="81ab6-158">取消注册的功能本身是异步的。</span><span class="sxs-lookup"><span data-stu-id="81ab6-158">The deregister function is itself asynchronous.</span></span> <span data-ttu-id="81ab6-159">因此，如果您有不应在注销完成之后运行的代码，则必须使用 `await` 关键字或 `then` 方法（如以下示例中所示）来等待取消注册功能。</span><span class="sxs-lookup"><span data-stu-id="81ab6-159">So, if you have code that should not run until after the deregistration is complete, then the deregister function should also be awaited with either the `await` keyword or with a `then` method as in the following examples.</span></span>

<span data-ttu-id="81ab6-160">取消注册处理程序：</span><span class="sxs-lookup"><span data-stu-id="81ab6-160">To deregister the handler:</span></span>

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```

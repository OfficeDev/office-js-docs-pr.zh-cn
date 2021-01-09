---
title: 显示或隐藏 Office 加载项的任务窗格
description: 了解如何在加载项连续运行时以编程方式隐藏或显示该加载项的用户界面。
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 20db609a3a6ded5624391f705dab1ad6b8f6e043
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789218"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a><span data-ttu-id="701f1-103">显示或隐藏 Office 加载项的任务窗格</span><span class="sxs-lookup"><span data-stu-id="701f1-103">Show or hide the task pane of your Office Add-in</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="701f1-104">可以通过调用函数来显示 Office 外接程序的任务 `Office.addin.showAsTaskpane()` 窗格。</span><span class="sxs-lookup"><span data-stu-id="701f1-104">You can show the task pane of your Office Add-in by calling the `Office.addin.showAsTaskpane()` function.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="701f1-105">前面的代码假定存在一个名为 **CurrentQuarterSales** 的 Excel 工作表的方案。</span><span class="sxs-lookup"><span data-stu-id="701f1-105">The previous code assumes a scenario where there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="701f1-106">只要激活此工作表，加载项就会使任务窗格可见。</span><span class="sxs-lookup"><span data-stu-id="701f1-106">The add-in will make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="701f1-107">该方法 `onCurrentQuarter` 是已注册工作表的 [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) 事件的处理程序。</span><span class="sxs-lookup"><span data-stu-id="701f1-107">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) event which has been registered for the worksheet.</span></span>

<span data-ttu-id="701f1-108">您还可以通过调用函数隐藏任务 `Office.addin.hide()` 窗格。</span><span class="sxs-lookup"><span data-stu-id="701f1-108">You can also hide the task pane by calling the `Office.addin.hide()` function.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

<span data-ttu-id="701f1-109">前面的代码为 [Office.Worksheet.onDeactivated 事件注册的](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated) 处理程序。</span><span class="sxs-lookup"><span data-stu-id="701f1-109">The previous code is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated) event.</span></span>

## <a name="additional-details-on-showing-the-task-pane"></a><span data-ttu-id="701f1-110">有关显示任务窗格的其他详细信息</span><span class="sxs-lookup"><span data-stu-id="701f1-110">Additional details on showing the task pane</span></span>

<span data-ttu-id="701f1-111">调用时，Office 将在任务窗格中显示你分配为资源 ID 的文件 () `Office.addin.showAsTaskpane()` `resid` 任务窗格的值。</span><span class="sxs-lookup"><span data-stu-id="701f1-111">When you call `Office.addin.showAsTaskpane()`, Office will display in a task pane the file that you assigned as the resource ID (`resid`) value of the task pane.</span></span> <span data-ttu-id="701f1-112">`resid`此值可通过打开文件并位于元素manifest.xml来分配 `<SourceLocation>` 或 `<Action xsi:type="ShowTaskpane">` 更改。</span><span class="sxs-lookup"><span data-stu-id="701f1-112">This `resid` value can be assigned or changed by opening your **manifest.xml** file and locating `<SourceLocation>` inside the `<Action xsi:type="ShowTaskpane">` element.</span></span>
<span data-ttu-id="701f1-113"> (有关其他详细信息， [请参阅"将 Office 外接程序](configure-your-add-in-to-use-a-shared-runtime.md) 配置为使用共享运行时"。) </span><span class="sxs-lookup"><span data-stu-id="701f1-113">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md) for additional details.)</span></span>

<span data-ttu-id="701f1-114">由于 `Office.addin.showAsTaskpane()` 是异步方法，因此代码将继续运行，直到函数完成。</span><span class="sxs-lookup"><span data-stu-id="701f1-114">Since `Office.addin.showAsTaskpane()` is an asynchronous method, your code will continue running until the function is complete.</span></span> <span data-ttu-id="701f1-115">使用关键字或方法等待完成， `await` `then()` 具体取决于你使用的 JavaScript 语法。</span><span class="sxs-lookup"><span data-stu-id="701f1-115">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span>

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a><span data-ttu-id="701f1-116">将加载项配置为使用共享运行时</span><span class="sxs-lookup"><span data-stu-id="701f1-116">Configure your add-in to use the shared runtime</span></span>

<span data-ttu-id="701f1-117">若要使用 `showAsTaskpane()` `hide()` 方法和方法，加载项必须使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="701f1-117">To use the `showAsTaskpane()` and `hide()` methods, your add-in must use the shared runtime.</span></span> <span data-ttu-id="701f1-118">有关详细信息，请参阅配置 [Office 外接程序以使用共享运行时](configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="701f1-118">For more information, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="701f1-119">状态和事件侦听器的保留</span><span class="sxs-lookup"><span data-stu-id="701f1-119">Preservation of state and event listeners</span></span>

<span data-ttu-id="701f1-120">方法和 `hide()` `showAsTaskpane()` 方法仅更改 *任务* 窗格的可见性。</span><span class="sxs-lookup"><span data-stu-id="701f1-120">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="701f1-121">它们不会卸载或重新加载它 (或重新初始化其状态) 。</span><span class="sxs-lookup"><span data-stu-id="701f1-121">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="701f1-122">请考虑以下方案：使用选项卡设计任务窗格。</span><span class="sxs-lookup"><span data-stu-id="701f1-122">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="701f1-123">首次 **启动** 加载项时，"主页"选项卡将打开。</span><span class="sxs-lookup"><span data-stu-id="701f1-123">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="701f1-124">假设用户打开"设置 **"** 选项卡，稍后任务窗格中的代码将调用 `hide()` 以响应某些事件。</span><span class="sxs-lookup"><span data-stu-id="701f1-124">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="701f1-125">稍后代码调用 `showAsTaskpane()` 以响应另一个事件。</span><span class="sxs-lookup"><span data-stu-id="701f1-125">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="701f1-126">任务窗格将重新出现，并且"设置 **"** 选项卡仍处于选中状态。</span><span class="sxs-lookup"><span data-stu-id="701f1-126">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![任务窗格的屏幕截图，其中四个选项卡标有"主页、设置、收藏夹和帐户"。](../images/TaskpaneWithTabs.png)

<span data-ttu-id="701f1-128">此外，在任务窗格中注册的任何事件侦听器将继续运行，即使任务窗格处于隐藏状态。</span><span class="sxs-lookup"><span data-stu-id="701f1-128">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="701f1-129">请考虑以下方案：任务窗格具有 Excel 的注册处理程序和名为 `Worksheet.onActivated` `Worksheet.onDeactivated` **Sheet1 的工作表的事件**。</span><span class="sxs-lookup"><span data-stu-id="701f1-129">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="701f1-130">激活的处理程序导致任务窗格中出现一个绿色点。</span><span class="sxs-lookup"><span data-stu-id="701f1-130">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="701f1-131">停用的处理程序将点红色 (，这是其默认状态) 。</span><span class="sxs-lookup"><span data-stu-id="701f1-131">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="701f1-132">假设代码在 `hide()` **Sheet1** 未激活且点为红色时调用。</span><span class="sxs-lookup"><span data-stu-id="701f1-132">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="701f1-133">隐藏任务窗格时，**将激活 Sheet1。**</span><span class="sxs-lookup"><span data-stu-id="701f1-133">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="701f1-134">稍后代码调用 `showAsTaskpane()` 以响应某些事件。</span><span class="sxs-lookup"><span data-stu-id="701f1-134">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="701f1-135">任务窗格打开时，该点为绿色，因为即使任务窗格处于隐藏状态，事件侦听器和处理程序也运行。</span><span class="sxs-lookup"><span data-stu-id="701f1-135">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

## <a name="handle-the-visibility-changed-event"></a><span data-ttu-id="701f1-136">处理可见性更改事件</span><span class="sxs-lookup"><span data-stu-id="701f1-136">Handle the visibility changed event</span></span>

<span data-ttu-id="701f1-137">当代码更改任务窗格的可见性时 `showAsTaskpane()` ，Office 将 `hide()` 触发 `VisibilityModeChanged` 该事件。</span><span class="sxs-lookup"><span data-stu-id="701f1-137">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="701f1-138">处理此事件可能很有用。</span><span class="sxs-lookup"><span data-stu-id="701f1-138">It can be useful to handle this event.</span></span> <span data-ttu-id="701f1-139">例如，假设任务窗格显示工作簿中所有工作表的列表。</span><span class="sxs-lookup"><span data-stu-id="701f1-139">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="701f1-140">如果在隐藏任务窗格时添加新工作表，则使任务窗格可见本身不会将新的工作表名称添加到列表中。</span><span class="sxs-lookup"><span data-stu-id="701f1-140">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="701f1-141">但代码可以响应该事件，以 `VisibilityModeChanged` 重新加载[workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) [Worksheet.name](/javascript/api/excel/excel.worksheet#name)中所有工作表的 Worksheet.name 属性，如下面的示例代码所示。</span><span class="sxs-lookup"><span data-stu-id="701f1-141">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="701f1-142">若要为事件注册处理程序，请不要像在大多数 Office JavaScript 上下文中一样使用"add handler"方法。</span><span class="sxs-lookup"><span data-stu-id="701f1-142">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="701f1-143">相反，有一个特殊的函数要传递给处理程序 [：Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-)。</span><span class="sxs-lookup"><span data-stu-id="701f1-143">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="701f1-144">示例如下。</span><span class="sxs-lookup"><span data-stu-id="701f1-144">The following is an example.</span></span> <span data-ttu-id="701f1-145">请注意，该属性 `args.visibilityMode` 的类型为 [VisibilityMode](/javascript/api/office/office.visibilitymode)。</span><span class="sxs-lookup"><span data-stu-id="701f1-145">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="701f1-146">该函数返回另一个取消 *注册处理程序* 的函数。</span><span class="sxs-lookup"><span data-stu-id="701f1-146">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="701f1-147">下面是一个简单但不稳固的示例：</span><span class="sxs-lookup"><span data-stu-id="701f1-147">Here is a simple, but not robust, example:</span></span>

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="701f1-148">此方法 `onVisibilityModeChanged` 是异步的，并返回一个承诺，这意味着代码需要等待承诺的实现，然后才能调用取消注册处理程序。 </span><span class="sxs-lookup"><span data-stu-id="701f1-148">The `onVisibilityModeChanged` method is asynchronous and returns a promise, which means that your code needs to await the fulfillment of the promise before it can call the **deregister** handler.</span></span>

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="701f1-149">取消注册函数也是异步的，并返回一个承诺。</span><span class="sxs-lookup"><span data-stu-id="701f1-149">The deregister function is also asynchronous and returns a promise.</span></span> <span data-ttu-id="701f1-150">因此，如果您有在取消注册完成之前不应运行的代码，则应该等待取消注册函数返回的承诺。</span><span class="sxs-lookup"><span data-stu-id="701f1-150">So, if you have code that should not run until after the deregistration is complete, then you should await the promise returned by the deregister function.</span></span>

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a><span data-ttu-id="701f1-151">另请参阅</span><span class="sxs-lookup"><span data-stu-id="701f1-151">See also</span></span>

- [<span data-ttu-id="701f1-152">将 Office 外接程序配置为使用共享的 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="701f1-152">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="701f1-153">文档打开时在 Office 外接程序中运行代码</span><span class="sxs-lookup"><span data-stu-id="701f1-153">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)

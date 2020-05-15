---
title: 启用和禁用加载项命令
description: 了解如何更改 Office Web 加载项中的自定义功能区按钮和菜单项的启用或禁用状态。
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: fa4830c0112486bbad7a13edf78e0c8c4277e143
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217891"
---
# <a name="enable-and-disable-add-in-commands"></a><span data-ttu-id="e5e70-103">启用和禁用加载项命令</span><span class="sxs-lookup"><span data-stu-id="e5e70-103">Enable and Disable Add-in Commands</span></span>

<span data-ttu-id="e5e70-104">如果加载项中的某些功能应仅适用于某些上下文，则能够以编程方式启用或禁用自定义加载项命令。</span><span class="sxs-lookup"><span data-stu-id="e5e70-104">When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands.</span></span> <span data-ttu-id="e5e70-105">例如，仅当光标位于表格中时，才启用用于更改表格标题的函数。</span><span class="sxs-lookup"><span data-stu-id="e5e70-105">For example, a function that changes the header of a table should only be enabled when the cursor is in a table.</span></span>

<span data-ttu-id="e5e70-106">你还可以指定 Office 主机应用程序打开时是启用还是禁用命令。</span><span class="sxs-lookup"><span data-stu-id="e5e70-106">You can also specify whether the command is enabled or disabled when the Office host application opens.</span></span>

> [!NOTE]
> <span data-ttu-id="e5e70-107">本文假定你熟悉以下文档。</span><span class="sxs-lookup"><span data-stu-id="e5e70-107">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="e5e70-108">如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。</span><span class="sxs-lookup"><span data-stu-id="e5e70-108">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> [<span data-ttu-id="e5e70-109">加载项命令的基本概念</span><span class="sxs-lookup"><span data-stu-id="e5e70-109">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

## <a name="rules-and-gotchas"></a><span data-ttu-id="e5e70-110">规则和陷阱</span><span class="sxs-lookup"><span data-stu-id="e5e70-110">Rules and gotchas</span></span>

### <a name="single-line-ribbon-in-office-on-the-web"></a><span data-ttu-id="e5e70-111">Office 网页版中的单行功能区</span><span class="sxs-lookup"><span data-stu-id="e5e70-111">Single-line ribbon in Office on the web</span></span>

<span data-ttu-id="e5e70-112">在 Office 网页版中，本文介绍的 API 和清单标记仅影响单行功能区。</span><span class="sxs-lookup"><span data-stu-id="e5e70-112">In Office on the web, the APIs and manifest markup described in this article only affect the single-line ribbon.</span></span> <span data-ttu-id="e5e70-113">它们不会对多行功能区产生任何影响。</span><span class="sxs-lookup"><span data-stu-id="e5e70-113">They have no effect on the multiline ribbon.</span></span> <span data-ttu-id="e5e70-114">它们会影响 Office 桌面版的这两个功能区。</span><span class="sxs-lookup"><span data-stu-id="e5e70-114">They affect both ribbons for desktop Office.</span></span> <span data-ttu-id="e5e70-115">有关这两个功能区的详细信息，请参阅[使用简化功能区](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98)。</span><span class="sxs-lookup"><span data-stu-id="e5e70-115">For more information about the two ribbons, see [Use the simplified ribbon](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98).</span></span>

### <a name="shared-runtime-required"></a><span data-ttu-id="e5e70-116">需要共享运行时</span><span class="sxs-lookup"><span data-stu-id="e5e70-116">Shared runtime required</span></span>

<span data-ttu-id="e5e70-117">本文介绍的 API 和清单标记，需要加载项清单指定它们应使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="e5e70-117">The APIs and manifest markup described in this article require that the add-in's manifest specify that it should use a shared runtime.</span></span> <span data-ttu-id="e5e70-118">为此，请执行下列步骤。</span><span class="sxs-lookup"><span data-stu-id="e5e70-118">To do this take the following steps.</span></span>

1. <span data-ttu-id="e5e70-119">在清单中的 [Runtimes](../reference/manifest/runtimes.md) 元素中，添加以下子元素：`<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`。</span><span class="sxs-lookup"><span data-stu-id="e5e70-119">In the [Runtimes](../reference/manifest/runtimes.md) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span></span> <span data-ttu-id="e5e70-120">（如果清单中尚无 `<Runtimes>` 元素，请在 `VersionOverrides` 部分中的 `<Host>` 元素下将其创建为第一个子元素。）</span><span class="sxs-lookup"><span data-stu-id="e5e70-120">(If there isn't already a `<Runtimes>` element in the manifest, create it as the first child under the `<Host>` element in the `VersionOverrides` section.)</span></span>
2. <span data-ttu-id="e5e70-121">在清单的 [Resources.Urls](../reference/manifest/resources.md) 部分中，添加以下子元素：`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`，其中 `{MyDomain}` 是加载项的域，`{path-to-start-page}` 是加载项的起始页路径；例如，`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`。</span><span class="sxs-lookup"><span data-stu-id="e5e70-121">In the [Resources.Urls](../reference/manifest/resources.md) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span></span>
3. <span data-ttu-id="e5e70-122">根据你的加载项是包含任务窗格、函数文件还是 Excel 自定义函数，你必须执行以下三个步骤中的一个或多个步骤：</span><span class="sxs-lookup"><span data-stu-id="e5e70-122">Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps:</span></span>

    - <span data-ttu-id="e5e70-123">如果加载项包含任务窗格，请将 [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) 元素的 `resid` 属性设置为与步骤 1 所使用 `<Runtime>` 元素的 `resid` 相同的字符串，例如 `Contoso.SharedRuntime.Url`。</span><span class="sxs-lookup"><span data-stu-id="e5e70-123">If the add-in contains a task pane, set the `resid` attribute of the [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="e5e70-124">该元素应如下所示：`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`。</span><span class="sxs-lookup"><span data-stu-id="e5e70-124">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="e5e70-125">如果加载项包含 Excel 自定义函数，请将 [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) 元素的 `resid` 属性设置为与步骤 1 所使用 `<Runtime>` 元素的 `resid` 相同的字符串，例如 `Contoso.SharedRuntime.Url`。</span><span class="sxs-lookup"><span data-stu-id="e5e70-125">If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) element exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="e5e70-126">该元素应如下所示：`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`。</span><span class="sxs-lookup"><span data-stu-id="e5e70-126">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="e5e70-127">如果加载项包含函数文件，请将 [FunctionFile](../reference/manifest/functionfile.md) 元素的 `resid` 属性设置为与步骤 1 所使用 `<Runtime>` 元素的 `resid` 相同的字符串，例如 `Contoso.SharedRuntime.Url`。</span><span class="sxs-lookup"><span data-stu-id="e5e70-127">If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](../reference/manifest/functionfile.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="e5e70-128">该元素应如下所示：`<FunctionFile resid="Contoso.SharedRuntime.Url"/>`。</span><span class="sxs-lookup"><span data-stu-id="e5e70-128">The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span></span>

## <a name="set-the-default-state-to-disabled"></a><span data-ttu-id="e5e70-129">将默认状态设置为“已禁用”</span><span class="sxs-lookup"><span data-stu-id="e5e70-129">Set the default state to disabled</span></span>

<span data-ttu-id="e5e70-130">默认情况下，当 Office 应用程序启动时，将启用任何加载项命令。</span><span class="sxs-lookup"><span data-stu-id="e5e70-130">By default, any Add-in Command is enabled when the Office application launches.</span></span> <span data-ttu-id="e5e70-131">如果要在 Office 应用程序启动时禁用自定义按钮或菜单项，请在清单中指定它。</span><span class="sxs-lookup"><span data-stu-id="e5e70-131">If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest.</span></span> <span data-ttu-id="e5e70-132">只需在控件的声明中的 [Action](../reference/manifest/action.md) 元素的*下方*（不在内部）之后立即添加 [Enabled ](../reference/manifest/enabled.md)元素（值为 `false`）即可。</span><span class="sxs-lookup"><span data-stu-id="e5e70-132">Just add an [Enabled](../reference/manifest/enabled.md) element (with the value `false`) immediately *below* (not inside) the [Action](../reference/manifest/action.md) element in the declaration of the control.</span></span> <span data-ttu-id="e5e70-133">下面显示了基本结构：</span><span class="sxs-lookup"><span data-stu-id="e5e70-133">The following shows the basic structure:</span></span>

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a><span data-ttu-id="e5e70-134">以编程方式更改状态</span><span class="sxs-lookup"><span data-stu-id="e5e70-134">Change the state programmatically</span></span>

<span data-ttu-id="e5e70-135">更改加载项命令的启用状态的基本步骤如下：</span><span class="sxs-lookup"><span data-stu-id="e5e70-135">The essential steps to changing the enabled status of an Add-in Command are:</span></span>

1. <span data-ttu-id="e5e70-136">创建 [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) 对象，该对象 (1) 按清单中指定的 ID 来指定命令及其父选项卡；以及 (2) 指定命令的启用或禁用状态。</span><span class="sxs-lookup"><span data-stu-id="e5e70-136">Create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the command, and its parent tab, by their IDs as specified in the manifest; and (2) specifies the enabled or disabled state of the command.</span></span>
2. <span data-ttu-id="e5e70-137">将 **RibbonUpdaterData** 对象传递到 [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-) 方法。</span><span class="sxs-lookup"><span data-stu-id="e5e70-137">Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-) method.</span></span>

<span data-ttu-id="e5e70-138">下面展示了一个非常简单的示例。</span><span class="sxs-lookup"><span data-stu-id="e5e70-138">The following is a simple example.</span></span> <span data-ttu-id="e5e70-139">请注意，“MyButton”和“OfficeAddinTab1”是从清单中复制的。</span><span class="sxs-lookup"><span data-stu-id="e5e70-139">Note that "MyButton" and "OfficeAddinTab1" are copied from the manifest.</span></span>

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "OfficeAppTab1", 
                controls: [
                {
                    id: "MyButton", 
                    enabled: true
                }
            ]}
        ]});
}
```

<span data-ttu-id="e5e70-140">我们还提供了几个接口（类型），使构建 **RibbonUpdateData** 对象变得更加容易。</span><span class="sxs-lookup"><span data-stu-id="e5e70-140">We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.</span></span> <span data-ttu-id="e5e70-141">下面是 TypeScript 中的等效示例，它利用了这些类型。</span><span class="sxs-lookup"><span data-stu-id="e5e70-141">The following is the equivalent example in TypeScript and it makes use of these types.</span></span>

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentTab: Tab = {id: "OfficeAddinTab1", controls: [button]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="e5e70-142">Office 控制何时更新功能区的状态。</span><span class="sxs-lookup"><span data-stu-id="e5e70-142">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="e5e70-143">**requestUpdate()** 方法会将更新请求加入队列中。</span><span class="sxs-lookup"><span data-stu-id="e5e70-143">The **requestUpdate()** method queues a request to update.</span></span> <span data-ttu-id="e5e70-144">在将请求加入队列后，该方法会立即解析 Promise 对象，而不是在功能区实际更新时解析。</span><span class="sxs-lookup"><span data-stu-id="e5e70-144">The method will resolve the Promise object as soon as it has queued the request, not when the ribbon actually updates.</span></span>

## <a name="change-the-state-in-response-to-an-event"></a><span data-ttu-id="e5e70-145">更改状态以响应事件</span><span class="sxs-lookup"><span data-stu-id="e5e70-145">Change the state in response to an event</span></span>

<span data-ttu-id="e5e70-146">一种应更改功能区状态的常见场景是用户启动的事件更改加载项上下文时。</span><span class="sxs-lookup"><span data-stu-id="e5e70-146">A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.</span></span>

<span data-ttu-id="e5e70-147">考虑这样一种场景：当且仅当激活图表时，才应启用按钮。</span><span class="sxs-lookup"><span data-stu-id="e5e70-147">Consider a scenario in which a button should be enabled when, and only when, a chart is activated.</span></span> <span data-ttu-id="e5e70-148">第一步是将清单中按钮的 [Enabled](../reference/manifest/enabled.md) 元素设置为 `false`。</span><span class="sxs-lookup"><span data-stu-id="e5e70-148">The first step is to set the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest to `false`.</span></span> <span data-ttu-id="e5e70-149">请参阅上面的示例。</span><span class="sxs-lookup"><span data-stu-id="e5e70-149">See above for an example.</span></span>

<span data-ttu-id="e5e70-150">第二步是分配处理程序。</span><span class="sxs-lookup"><span data-stu-id="e5e70-150">Second, assign handlers.</span></span> <span data-ttu-id="e5e70-151">这通常在 **Office.onReady** 方法中完成，如以下示例所示，该示例将处理程序（在后续步骤中创建）分配给工作表中所有图表的 **onActivated** 和 **onDeactivated** 事件。</span><span class="sxs-lookup"><span data-stu-id="e5e70-151">This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

<span data-ttu-id="e5e70-152">第三步是定义 `enableChartFormat` 处理程序。</span><span class="sxs-lookup"><span data-stu-id="e5e70-152">Third, define the `enableChartFormat` handler.</span></span> <span data-ttu-id="e5e70-153">以下是一个简单示例，请参阅下面的**最佳做法：测试控件状态错误**，以获取更改控件状态的更可靠方法。</span><span class="sxs-lookup"><span data-stu-id="e5e70-153">The following is a simple example, but see **Best practice: Test for control status errors** below for a more robust way of changing a control's status.</span></span>

```javascript
function enableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: true};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="e5e70-154">第四步是定义 `disableChartFormat` 处理程序。</span><span class="sxs-lookup"><span data-stu-id="e5e70-154">Fourth, define the `disableChartFormat` handler.</span></span> <span data-ttu-id="e5e70-155">除了将按钮对象的 **enabled** 属性设置为 `false` 之外，其他操作与 `enableChartFormat` 相同。</span><span class="sxs-lookup"><span data-stu-id="e5e70-155">It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.</span></span>

## <a name="best-practice-test-for-control-status-errors"></a><span data-ttu-id="e5e70-156">最佳做法：测试控件状态错误</span><span class="sxs-lookup"><span data-stu-id="e5e70-156">Best practice: Test for control status errors</span></span>

<span data-ttu-id="e5e70-157">在某些情况下，调用 `requestUpdate` 后，功能区不会重画，因此控件的可单击状态不会发生更改。</span><span class="sxs-lookup"><span data-stu-id="e5e70-157">In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change.</span></span> <span data-ttu-id="e5e70-158">因此，加载项的最佳做法是跟踪其控件的状态。</span><span class="sxs-lookup"><span data-stu-id="e5e70-158">For this reason it is a best practice for the add-in to keep track of the status of its controls.</span></span> <span data-ttu-id="e5e70-159">加载项应符合以下规则：</span><span class="sxs-lookup"><span data-stu-id="e5e70-159">The add-in should conform to these rules:</span></span>

1. <span data-ttu-id="e5e70-160">每当调用 `requestUpdate` 时，代码都应记录自定义按钮和菜单项的预期状态。</span><span class="sxs-lookup"><span data-stu-id="e5e70-160">Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.</span></span>
2. <span data-ttu-id="e5e70-161">单击自定义控件时，处理程序中的第一个代码应检查该按钮是否应为可单击按钮。</span><span class="sxs-lookup"><span data-stu-id="e5e70-161">When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable.</span></span> <span data-ttu-id="e5e70-162">如果不是，则该代码应报告或记录错误，然后再次尝试将按钮设置为预期状态。</span><span class="sxs-lookup"><span data-stu-id="e5e70-162">If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.</span></span>

<span data-ttu-id="e5e70-163">以下示例显示用于禁用按钮和记录按钮状态的函数。</span><span class="sxs-lookup"><span data-stu-id="e5e70-163">The following example shows a function that disables a button and records the button's status.</span></span> <span data-ttu-id="e5e70-164">请注意，`chartFormatButtonEnabled` 是全局布尔变量，其初始化为与清单中按钮的 [Enabled](../reference/manifest/enabled.md) 元素相同的值。</span><span class="sxs-lookup"><span data-stu-id="e5e70-164">Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.</span></span>

```javascript
function disableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: false};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

<span data-ttu-id="e5e70-165">以下示例显示按钮的处理程序如何测试按钮的错误状态。</span><span class="sxs-lookup"><span data-stu-id="e5e70-165">The following example shows how the button's handler tests for an incorrect state of the button.</span></span> <span data-ttu-id="e5e70-166">请注意，`reportError` 是用于显示或记录错误的函数。</span><span class="sxs-lookup"><span data-stu-id="e5e70-166">Note that `reportError` is a function that shows or logs an error.</span></span>

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here

    } else {
        // Report the error and try again to disable.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## <a name="error-handling"></a><span data-ttu-id="e5e70-167">错误处理</span><span class="sxs-lookup"><span data-stu-id="e5e70-167">Error handling</span></span>

<span data-ttu-id="e5e70-168">在某些情况下，Office 无法更新功能区，并将返回错误。</span><span class="sxs-lookup"><span data-stu-id="e5e70-168">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="e5e70-169">例如，如果升级了加载项，并且升级后的加载项具有一组不同的自定义加载项命令，则必须关闭并重新打开 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="e5e70-169">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="e5e70-170">在此之前，`requestUpdate` 方法将返回错误 `HostRestartNeeded`。</span><span class="sxs-lookup"><span data-stu-id="e5e70-170">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="e5e70-171">以下是如何处理此错误的示例。</span><span class="sxs-lookup"><span data-stu-id="e5e70-171">The following is an example of how to handle this error.</span></span> <span data-ttu-id="e5e70-172">在此示例中，`reportError` 方法向用户显示错误。</span><span class="sxs-lookup"><span data-stu-id="e5e70-172">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function disableChartFormat() {
    try {
        var button = {id: "ChartFormatButton", enabled: false};
        var parentTab = {id: "CustomChartTab", controls: [button]};
        var ribbonUpdater = {tabs: [parentTab]};
        await Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```

## <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="e5e70-173">使用要求集测试平台支持</span><span class="sxs-lookup"><span data-stu-id="e5e70-173">Test for platform support with requirement sets</span></span>

<span data-ttu-id="e5e70-p122">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="e5e70-p122">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="e5e70-177">启用/禁用 API 需要支持以下要求集：</span><span class="sxs-lookup"><span data-stu-id="e5e70-177">The enable/disable APIs require support of the following requirement set:</span></span>

- [<span data-ttu-id="e5e70-178">AddinCommands 1.1</span><span class="sxs-lookup"><span data-stu-id="e5e70-178">AddinCommands 1.1</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)

---
title: 启用和禁用加载项命令
description: 了解如何更改 Office Web 加载项中的自定义功能区按钮和菜单项的启用或禁用状态。
ms.date: 04/30/2021
localization_priority: Normal
ms.openlocfilehash: 9690850b2206c09b99dfc826dae1ecef915d5a04
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330155"
---
# <a name="enable-and-disable-add-in-commands"></a><span data-ttu-id="deaf0-103">启用和禁用加载项命令</span><span class="sxs-lookup"><span data-stu-id="deaf0-103">Enable and Disable Add-in Commands</span></span>

<span data-ttu-id="deaf0-104">如果加载项中的某些功能应仅适用于某些上下文，则能够以编程方式启用或禁用自定义加载项命令。</span><span class="sxs-lookup"><span data-stu-id="deaf0-104">When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands.</span></span> <span data-ttu-id="deaf0-105">例如，仅当光标位于表格中时，才启用用于更改表格标题的函数。</span><span class="sxs-lookup"><span data-stu-id="deaf0-105">For example, a function that changes the header of a table should only be enabled when the cursor is in a table.</span></span>

<span data-ttu-id="deaf0-106">还可以指定在客户端应用程序打开时是启用还是禁用Office命令。</span><span class="sxs-lookup"><span data-stu-id="deaf0-106">You can also specify whether the command is enabled or disabled when the Office client application opens.</span></span>

> [!NOTE]
> <span data-ttu-id="deaf0-107">本文假定你熟悉以下文档。</span><span class="sxs-lookup"><span data-stu-id="deaf0-107">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="deaf0-108">如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。</span><span class="sxs-lookup"><span data-stu-id="deaf0-108">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="deaf0-109">加载项命令的基本概念</span><span class="sxs-lookup"><span data-stu-id="deaf0-109">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a><span data-ttu-id="deaf0-110">Office应用程序和平台仅支持</span><span class="sxs-lookup"><span data-stu-id="deaf0-110">Office application and platform support only</span></span>

<span data-ttu-id="deaf0-111">本文中介绍的 API 仅可用于Excel平台和 PowerPoint web 版。</span><span class="sxs-lookup"><span data-stu-id="deaf0-111">The APIs described in this article are only available in Excel on all platforms and in PowerPoint on the web.</span></span>

### <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="deaf0-112">使用要求集测试平台支持</span><span class="sxs-lookup"><span data-stu-id="deaf0-112">Test for platform support with requirement sets</span></span>

<span data-ttu-id="deaf0-113">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="deaf0-113">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="deaf0-114">Office外接程序使用清单中指定的要求集或使用运行时检查来确定 Office 应用程序和平台组合是否支持外接程序所需的 API。</span><span class="sxs-lookup"><span data-stu-id="deaf0-114">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application and platform combination supports APIs that an add-in needs.</span></span> <span data-ttu-id="deaf0-115">有关详细信息，请参阅Office[版本和要求集](../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="deaf0-115">For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="deaf0-116">启用/禁用 API 属于 [RibbonApi 1.1](../reference/requirement-sets/ribbon-api-requirement-sets.md) 要求集。</span><span class="sxs-lookup"><span data-stu-id="deaf0-116">The enable/disable APIs belong to the [RibbonApi 1.1](../reference/requirement-sets/ribbon-api-requirement-sets.md) requirement set.</span></span>

> [!NOTE]
> <span data-ttu-id="deaf0-117">**RibbonApi 1.1** 要求集在清单中尚不受支持，因此您无法在清单的 部分中指定 `<Requirements>` 它。</span><span class="sxs-lookup"><span data-stu-id="deaf0-117">The **RibbonApi 1.1** requirement set is not yet supported in the manifest, so you cannot specify it in the manifest's `<Requirements>` section.</span></span> <span data-ttu-id="deaf0-118">若要测试支持，代码应调用 `Office.context.requirements.isSetSupported('RibbonApi', '1.1')` 。</span><span class="sxs-lookup"><span data-stu-id="deaf0-118">To test for support, your code should call `Office.context.requirements.isSetSupported('RibbonApi', '1.1')`.</span></span> <span data-ttu-id="deaf0-119">如果 *且仅在 返回* 时 ， `true` 代码可以调用启用/禁用 API。</span><span class="sxs-lookup"><span data-stu-id="deaf0-119">If, *and only if*, that call returns `true`, your code can call the enable/disable APIs.</span></span> <span data-ttu-id="deaf0-120">如果 调用 `isSetSupported` 返回 `false` ，则所有自定义外接程序命令将一向启用。</span><span class="sxs-lookup"><span data-stu-id="deaf0-120">If the call of `isSetSupported` returns `false`, then all custom add-in commands are enabled all of the time.</span></span> <span data-ttu-id="deaf0-121">您必须设计生产外接程序以及任何应用内说明，以考虑当 **RibbonApi 1.1** 要求集不受支持时它如何工作。</span><span class="sxs-lookup"><span data-stu-id="deaf0-121">You must design your production add-in, and any in-app instructions, to take account of how it will work when the **RibbonApi 1.1** requirement set is not supported.</span></span> <span data-ttu-id="deaf0-122">有关使用 有关详细信息和示例，请参阅指定 Office 应用程序和 API 要求， `isSetSupported` 尤其是在[JavaScript 代码中使用运行时检查](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)。 [](../develop/specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="deaf0-122">For more information and examples of using `isSetSupported`, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md), especially [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="deaf0-123"> (本文清单中的设置 [Requirements](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) 元素部分不适用于功能区 1.1.) </span><span class="sxs-lookup"><span data-stu-id="deaf0-123">(The section [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) of that article does not apply to Ribbon 1.1.)</span></span>

## <a name="shared-runtime-required"></a><span data-ttu-id="deaf0-124">需要共享运行时</span><span class="sxs-lookup"><span data-stu-id="deaf0-124">Shared runtime required</span></span>

<span data-ttu-id="deaf0-125">本文介绍的 API 和清单标记，需要加载项清单指定它们应使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="deaf0-125">The APIs and manifest markup described in this article require that the add-in's manifest specify that it should use a shared runtime.</span></span> <span data-ttu-id="deaf0-126">为此，请执行下列步骤。</span><span class="sxs-lookup"><span data-stu-id="deaf0-126">To do this take the following steps.</span></span>

1. <span data-ttu-id="deaf0-127">在清单中的 [Runtimes](../reference/manifest/runtimes.md) 元素中，添加以下子元素：`<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`。</span><span class="sxs-lookup"><span data-stu-id="deaf0-127">In the [Runtimes](../reference/manifest/runtimes.md) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span></span> <span data-ttu-id="deaf0-128">（如果清单中尚无 `<Runtimes>` 元素，请在 `VersionOverrides` 部分中的 `<Host>` 元素下将其创建为第一个子元素。）</span><span class="sxs-lookup"><span data-stu-id="deaf0-128">(If there isn't already a `<Runtimes>` element in the manifest, create it as the first child under the `<Host>` element in the `VersionOverrides` section.)</span></span>
2. <span data-ttu-id="deaf0-129">在清单的 [Resources.Urls](../reference/manifest/resources.md) 部分中，添加以下子元素：`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`，其中 `{MyDomain}` 是加载项的域，`{path-to-start-page}` 是加载项的起始页路径；例如，`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`。</span><span class="sxs-lookup"><span data-stu-id="deaf0-129">In the [Resources.Urls](../reference/manifest/resources.md) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span></span>
3. <span data-ttu-id="deaf0-130">根据你的加载项是包含任务窗格、函数文件还是 Excel 自定义函数，你必须执行以下三个步骤中的一个或多个步骤：</span><span class="sxs-lookup"><span data-stu-id="deaf0-130">Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps:</span></span>

    - <span data-ttu-id="deaf0-131">如果加载项包含任务窗格，请将 [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) 元素的 `resid` 属性设置为与步骤 1 所使用 `<Runtime>` 元素的 `resid` 相同的字符串，例如 `Contoso.SharedRuntime.Url`。</span><span class="sxs-lookup"><span data-stu-id="deaf0-131">If the add-in contains a task pane, set the `resid` attribute of the [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="deaf0-132">该元素应如下所示：`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`。</span><span class="sxs-lookup"><span data-stu-id="deaf0-132">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="deaf0-133">如果加载项包含 Excel 自定义函数，请将 [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) 元素的 `resid` 属性设置为与步骤 1 所使用 `<Runtime>` 元素的 `resid` 相同的字符串，例如 `Contoso.SharedRuntime.Url`。</span><span class="sxs-lookup"><span data-stu-id="deaf0-133">If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) element exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="deaf0-134">该元素应如下所示：`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`。</span><span class="sxs-lookup"><span data-stu-id="deaf0-134">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="deaf0-135">如果加载项包含函数文件，请将 [FunctionFile](../reference/manifest/functionfile.md) 元素的 `resid` 属性设置为与步骤 1 所使用 `<Runtime>` 元素的 `resid` 相同的字符串，例如 `Contoso.SharedRuntime.Url`。</span><span class="sxs-lookup"><span data-stu-id="deaf0-135">If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](../reference/manifest/functionfile.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="deaf0-136">该元素应如下所示：`<FunctionFile resid="Contoso.SharedRuntime.Url"/>`。</span><span class="sxs-lookup"><span data-stu-id="deaf0-136">The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span></span>

## <a name="set-the-default-state-to-disabled"></a><span data-ttu-id="deaf0-137">将默认状态设置为“已禁用”</span><span class="sxs-lookup"><span data-stu-id="deaf0-137">Set the default state to disabled</span></span>

<span data-ttu-id="deaf0-138">默认情况下，当 Office 应用程序启动时，将启用任何加载项命令。</span><span class="sxs-lookup"><span data-stu-id="deaf0-138">By default, any Add-in Command is enabled when the Office application launches.</span></span> <span data-ttu-id="deaf0-139">如果要在 Office 应用程序启动时禁用自定义按钮或菜单项，请在清单中指定它。</span><span class="sxs-lookup"><span data-stu-id="deaf0-139">If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest.</span></span> <span data-ttu-id="deaf0-140">只需在控件的声明中的 [Action](../reference/manifest/action.md) 元素的 *下方*（不在内部）之后立即添加 [Enabled](../reference/manifest/enabled.md)元素（值为 `false`）即可。</span><span class="sxs-lookup"><span data-stu-id="deaf0-140">Just add an [Enabled](../reference/manifest/enabled.md) element (with the value `false`) immediately *below* (not inside) the [Action](../reference/manifest/action.md) element in the declaration of the control.</span></span> <span data-ttu-id="deaf0-141">下面显示了基本结构：</span><span class="sxs-lookup"><span data-stu-id="deaf0-141">The following shows the basic structure:</span></span>

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a><span data-ttu-id="deaf0-142">以编程方式更改状态</span><span class="sxs-lookup"><span data-stu-id="deaf0-142">Change the state programmatically</span></span>

<span data-ttu-id="deaf0-143">更改加载项命令的启用状态的基本步骤如下：</span><span class="sxs-lookup"><span data-stu-id="deaf0-143">The essential steps to changing the enabled status of an Add-in Command are:</span></span>

1. <span data-ttu-id="deaf0-144">创建 [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) 对象 (1) 指定命令及其父组和选项卡，按清单中声明的其 ID;和 (2) 指定命令的启用或禁用状态。</span><span class="sxs-lookup"><span data-stu-id="deaf0-144">Create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the command, and its parent group and tab, by their IDs as declared in the manifest; and (2) specifies the enabled or disabled state of the command.</span></span>
2. <span data-ttu-id="deaf0-145">将 **RibbonUpdaterData** 对象传递到 [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) 方法。</span><span class="sxs-lookup"><span data-stu-id="deaf0-145">Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method.</span></span>

<span data-ttu-id="deaf0-146">下面展示了一个非常简单的示例。</span><span class="sxs-lookup"><span data-stu-id="deaf0-146">The following is a simple example.</span></span> <span data-ttu-id="deaf0-147">请注意，"MyButton"、"OfficeAddinTab1"和"CustomGroup111"从清单中复制。</span><span class="sxs-lookup"><span data-stu-id="deaf0-147">Note that "MyButton", "OfficeAddinTab1", and "CustomGroup111" are copied from the manifest.</span></span>

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "OfficeAppTab1", 
                groups: [
                    {
                      id: "CustomGroup111",
                      controls: [
                        {
                            id: "MyButton", 
                            enabled: true
                        }
                      ]
                    }
                ]
            }
        ]
    });
}
```

<span data-ttu-id="deaf0-148">我们还提供了几个接口（类型），使构建 **RibbonUpdateData** 对象变得更加容易。</span><span class="sxs-lookup"><span data-stu-id="deaf0-148">We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.</span></span> <span data-ttu-id="deaf0-149">下面是 TypeScript 中的等效示例，它利用了这些类型。</span><span class="sxs-lookup"><span data-stu-id="deaf0-149">The following is the equivalent example in TypeScript and it makes use of these types.</span></span>

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentGroup: Group = {id: "CustomGroup111", controls: [button]};
    const parentTab: Tab = {id: "OfficeAddinTab1", groups: [parentGroup]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="deaf0-150">如果父函数是异步 () 可以调用 `await` **requestUpdate，** 但请注意，Office应用程序在更新功能区的状态时进行控制。</span><span class="sxs-lookup"><span data-stu-id="deaf0-150">You can `await` the call of **requestUpdate()** if the parent function is asynchronous, but note that the Office application controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="deaf0-151">**requestUpdate()** 方法会将更新请求加入队列中。</span><span class="sxs-lookup"><span data-stu-id="deaf0-151">The **requestUpdate()** method queues a request to update.</span></span> <span data-ttu-id="deaf0-152">一旦将请求排入队列，该方法将解析承诺对象，而不是功能区实际更新时。</span><span class="sxs-lookup"><span data-stu-id="deaf0-152">The method will resolve the promise object as soon as it has queued the request, not when the ribbon actually updates.</span></span>

## <a name="change-the-state-in-response-to-an-event"></a><span data-ttu-id="deaf0-153">更改状态以响应事件</span><span class="sxs-lookup"><span data-stu-id="deaf0-153">Change the state in response to an event</span></span>

<span data-ttu-id="deaf0-154">一种应更改功能区状态的常见场景是用户启动的事件更改加载项上下文时。</span><span class="sxs-lookup"><span data-stu-id="deaf0-154">A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.</span></span>

<span data-ttu-id="deaf0-155">考虑这样一种场景：当且仅当激活图表时，才应启用按钮。</span><span class="sxs-lookup"><span data-stu-id="deaf0-155">Consider a scenario in which a button should be enabled when, and only when, a chart is activated.</span></span> <span data-ttu-id="deaf0-156">第一步是将清单中按钮的 [Enabled](../reference/manifest/enabled.md) 元素设置为 `false`。</span><span class="sxs-lookup"><span data-stu-id="deaf0-156">The first step is to set the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest to `false`.</span></span> <span data-ttu-id="deaf0-157">请参阅上面的示例。</span><span class="sxs-lookup"><span data-stu-id="deaf0-157">See above for an example.</span></span>

<span data-ttu-id="deaf0-158">第二步是分配处理程序。</span><span class="sxs-lookup"><span data-stu-id="deaf0-158">Second, assign handlers.</span></span> <span data-ttu-id="deaf0-159">这通常在 **Office.onReady** 方法中完成，如以下示例所示，该示例将处理程序（在后续步骤中创建）分配给工作表中所有图表的 **onActivated** 和 **onDeactivated** 事件。</span><span class="sxs-lookup"><span data-stu-id="deaf0-159">This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

<span data-ttu-id="deaf0-160">第三步是定义 `enableChartFormat` 处理程序。</span><span class="sxs-lookup"><span data-stu-id="deaf0-160">Third, define the `enableChartFormat` handler.</span></span> <span data-ttu-id="deaf0-161">以下是一个简单示例，请参阅下面的[最佳做法：测试控件状态错误](#best-practice-test-for-control-status-errors)，以获取更改控件状态的更可靠方法。</span><span class="sxs-lookup"><span data-stu-id="deaf0-161">The following is a simple example, but see [Best practice: Test for control status errors](#best-practice-test-for-control-status-errors) below for a more robust way of changing a control's status.</span></span>

```javascript
function enableChartFormat() {
    var button = {
                  id: "ChartFormatButton", 
                  enabled: true
                 };
    var parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    var parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    var ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="deaf0-162">第四步是定义 `disableChartFormat` 处理程序。</span><span class="sxs-lookup"><span data-stu-id="deaf0-162">Fourth, define the `disableChartFormat` handler.</span></span> <span data-ttu-id="deaf0-163">除了将按钮对象的 **enabled** 属性设置为 `false` 之外，其他操作与 `enableChartFormat` 相同。</span><span class="sxs-lookup"><span data-stu-id="deaf0-163">It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.</span></span>

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="deaf0-164">切换选项卡可见性和按钮的启用状态</span><span class="sxs-lookup"><span data-stu-id="deaf0-164">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="deaf0-165">**requestUpdate** 方法还用于切换自定义上下文选项卡的可见性。有关此代码和示例代码的详细信息，请参阅在加载项Office [上下文选项卡](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time)。</span><span class="sxs-lookup"><span data-stu-id="deaf0-165">The **requestUpdate** method is also used to toggle the visibility of a custom contextual tab. For details about this and example code, see [Create custom contextual tabs in Office Add-ins](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time).</span></span>

## <a name="best-practice-test-for-control-status-errors"></a><span data-ttu-id="deaf0-166">最佳做法：测试控件状态错误</span><span class="sxs-lookup"><span data-stu-id="deaf0-166">Best practice: Test for control status errors</span></span>

<span data-ttu-id="deaf0-167">在某些情况下，调用 `requestUpdate` 后，功能区不会重画，因此控件的可单击状态不会发生更改。</span><span class="sxs-lookup"><span data-stu-id="deaf0-167">In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change.</span></span> <span data-ttu-id="deaf0-168">因此，加载项的最佳做法是跟踪其控件的状态。</span><span class="sxs-lookup"><span data-stu-id="deaf0-168">For this reason it is a best practice for the add-in to keep track of the status of its controls.</span></span> <span data-ttu-id="deaf0-169">加载项应符合以下规则：</span><span class="sxs-lookup"><span data-stu-id="deaf0-169">The add-in should conform to these rules:</span></span>

1. <span data-ttu-id="deaf0-170">每当调用 `requestUpdate` 时，代码都应记录自定义按钮和菜单项的预期状态。</span><span class="sxs-lookup"><span data-stu-id="deaf0-170">Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.</span></span>
2. <span data-ttu-id="deaf0-171">单击自定义控件时，处理程序中的第一个代码应检查该按钮是否应为可单击按钮。</span><span class="sxs-lookup"><span data-stu-id="deaf0-171">When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable.</span></span> <span data-ttu-id="deaf0-172">如果不是，则该代码应报告或记录错误，然后再次尝试将按钮设置为预期状态。</span><span class="sxs-lookup"><span data-stu-id="deaf0-172">If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.</span></span>

<span data-ttu-id="deaf0-173">以下示例显示用于禁用按钮和记录按钮状态的函数。</span><span class="sxs-lookup"><span data-stu-id="deaf0-173">The following example shows a function that disables a button and records the button's status.</span></span> <span data-ttu-id="deaf0-174">请注意，`chartFormatButtonEnabled` 是全局布尔变量，其初始化为与清单中按钮的 [Enabled](../reference/manifest/enabled.md) 元素相同的值。</span><span class="sxs-lookup"><span data-stu-id="deaf0-174">Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.</span></span>

```javascript
function disableChartFormat() {
    var button = {
                  id: "ChartFormatButton", 
                  enabled: false
                 };
    var parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    var parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    var ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

<span data-ttu-id="deaf0-175">以下示例显示按钮的处理程序如何测试按钮的错误状态。</span><span class="sxs-lookup"><span data-stu-id="deaf0-175">The following example shows how the button's handler tests for an incorrect state of the button.</span></span> <span data-ttu-id="deaf0-176">请注意，`reportError` 是用于显示或记录错误的函数。</span><span class="sxs-lookup"><span data-stu-id="deaf0-176">Note that `reportError` is a function that shows or logs an error.</span></span>

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here

    } else {
        // Report the error and try again to disable.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## <a name="error-handling"></a><span data-ttu-id="deaf0-177">错误处理</span><span class="sxs-lookup"><span data-stu-id="deaf0-177">Error handling</span></span>

<span data-ttu-id="deaf0-178">在某些情况下，Office 无法更新功能区，并将返回错误。</span><span class="sxs-lookup"><span data-stu-id="deaf0-178">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="deaf0-179">例如，如果升级了加载项，并且升级后的加载项具有一组不同的自定义加载项命令，则必须关闭并重新打开 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="deaf0-179">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="deaf0-180">在此之前，`requestUpdate` 方法将返回错误 `HostRestartNeeded`。</span><span class="sxs-lookup"><span data-stu-id="deaf0-180">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="deaf0-181">以下是如何处理此错误的示例。</span><span class="sxs-lookup"><span data-stu-id="deaf0-181">The following is an example of how to handle this error.</span></span> <span data-ttu-id="deaf0-182">在此示例中，`reportError` 方法向用户显示错误。</span><span class="sxs-lookup"><span data-stu-id="deaf0-182">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function disableChartFormat() {
    try {
        var button = {
                      id: "ChartFormatButton", 
                      enabled: false
                     };
        var parentGroup = {
                           id: "MyGroup",
                           controls: [button]
                          };
        var parentTab = {
                         id: "CustomChartTab", 
                         groups: [parentGroup]
                        };
        var ribbonUpdater = {tabs: [parentTab]};
        Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```

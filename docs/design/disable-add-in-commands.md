---
title: 启用和禁用加载项命令
description: 了解如何更改 Office Web 加载项中的自定义功能区按钮和菜单项的启用或禁用状态。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 502c9247a6c63775c562dab7479e0ca926f14154
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423046"
---
# <a name="enable-and-disable-add-in-commands"></a>启用和禁用加载项命令

如果加载项中的某些功能应仅适用于某些上下文，则能够以编程方式启用或禁用自定义加载项命令。 例如，仅当光标位于表格中时，才启用用于更改表格标题的函数。

还可以指定在 Office 客户端应用程序打开时是否启用或禁用该命令。

> [!NOTE]
> 本文假定你熟悉以下文档。 如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。
>
> - [加载项命令的基本概念](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a>仅限 Office 应用程序和平台支持

本文中所述的 API 仅在 Excel、PowerPoint 和 Word 中可用。

### <a name="test-for-platform-support-with-requirement-sets"></a>使用要求集测试平台支持

要求集是指各组已命名的 API 成员。 Office 外接程序使用清单中指定的要求集，或使用运行时检查来确定 Office 应用程序和平台组合是否支持外接程序需要的 API。 有关详细信息，请参阅 [Office 版本和要求集](../develop/office-versions-and-requirement-sets.md)。

启用/禁用 API 属于 [RibbonApi 1.1](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets) 要求集。

> [!NOTE]
> 清单中尚不支持 **RibbonApi 1.1** 要求集，因此不能在清单部分 **\<Requirements\>** 中指定它。 若要测试支持，代码应调用 `Office.context.requirements.isSetSupported('RibbonApi', '1.1')`。 如果该调用返回， *并且仅当* 该调用返回 `true`时，代码可以调用启用/禁用 API。 如果调用 `isSetSupported` 返回 `false`，则所有自定义加载项命令将一直启用。 必须设计生产外接程序和任何应用内说明，以考虑在不支持 **RibbonApi 1.1** 要求集时其工作原理。 有关使用的 `isSetSupported`详细信息和示例，请参阅 [指定 Office 应用程序和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)，特别是 [运行时检查方法和要求集支持](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)。  (“ [指定哪些 Office 版本和平台可以托管该](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in) 文章的外接程序不适用于功能区 1.1.) 

## <a name="shared-runtime-required"></a>需要共享运行时

本文中所述的 API 和清单标记要求外接程序的清单指定它应使用 [共享运行时](../testing/runtimes.md#shared-runtime)。 为此，请执行以下步骤。

1. 在清单中的 [Runtimes](/javascript/api/manifest/runtimes) 元素中，添加以下子元素：`<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`。  (如果清单中还没有元素，请将其创建为 section.) 中元素 **\<VersionOverrides\>** 下 **\<Host\>** 的第一个 **\<Runtimes\>** 子元素。
2. 在清单的 [Resources.Urls](/javascript/api/manifest/resources) 部分中，添加以下子元素：`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`，其中 `{MyDomain}` 是加载项的域，`{path-to-start-page}` 是加载项的起始页路径；例如，`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`。
3. 根据外接程序是否包含任务窗格、函数文件或 Excel 自定义函数，必须执行以下三个步骤中的一个或多个步骤。

    - 如果加载项包含任务窗格，请设置 `resid` [Action](/javascript/api/manifest/action) 的属性。[SourceLocation](/javascript/api/manifest/sourcelocation) 元素与在步骤 1 中用于`resid`元素的 **\<Runtime\>** 字符串完全相同;例如。 `Contoso.SharedRuntime.Url` 该元素应如下所示：`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`。
    - 如果外接程序包含 Excel 自定义函数，请设置 `resid` [Page](/javascript/api/manifest/page) 的属性。[SourceLocation](/javascript/api/manifest/sourcelocation) 元素与在步骤 1 中用于`resid`元素的 **\<Runtime\>** 字符串完全相同;例如。 `Contoso.SharedRuntime.Url` 该元素应如下所示：`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`。
    - 如果外接程序包含函数文件，请将 [FunctionFile](/javascript/api/manifest/functionfile) 元素的属性设置`resid`为`resid`与步骤 1 中用于元素的 **\<Runtime\>** 字符串完全相同的字符串;例如。 `Contoso.SharedRuntime.Url` 该元素应如下所示：`<FunctionFile resid="Contoso.SharedRuntime.Url"/>`。

## <a name="set-the-default-state-to-disabled"></a>将默认状态设置为“已禁用”

默认情况下，当 Office 应用程序启动时，将启用任何加载项命令。 如果要在 Office 应用程序启动时禁用自定义按钮或菜单项，请在清单中指定它。 只需在控件的声明中的 [Action](/javascript/api/manifest/action) 元素的 *下方*（不在内部）之后立即添加 [Enabled](/javascript/api/manifest/enabled)元素（值为 `false`）即可。 下面显示了基本结构。

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
                <Control ... id="Contoso.MyButton3">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>以编程方式更改状态

更改加载项命令的启用状态的基本步骤如下：

1. 创建一个 [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) 对象，该对象 (1) 按清单中声明的命令及其父组和选项卡指定其 ID： (2) 指定命令的启用或禁用状态。
2. 将 **RibbonUpdaterData** 对象传递到 [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) 方法。

下面展示了一个非常简单的示例。 请注意，“MyButton”、“OfficeAddinTab1”和“CustomGroup111”是从清单中复制的。

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

我们还提供了几个接口（类型），使构建 **RibbonUpdateData** 对象变得更加容易。 下面是 TypeScript 中的等效示例，它利用了这些类型。

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentGroup: Group = {id: "CustomGroup111", controls: [button]};
    const parentTab: Tab = {id: "OfficeAddinTab1", groups: [parentGroup]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

如果父函数是异步的，则可以 `await` 调用 **requestUpdate ()** ，但请注意，Office 应用程序在更新功能区状态时会进行控制。 **requestUpdate()** 方法会将更新请求加入队列中。 该方法将在请求排队后立即解析 promise 对象，而不是当功能区实际更新时。

## <a name="change-the-state-in-response-to-an-event"></a>更改状态以响应事件

一种应更改功能区状态的常见场景是用户启动的事件更改加载项上下文时。

考虑这样一种场景：当且仅当激活图表时，才应启用按钮。 第一步是将清单中按钮的 [Enabled](/javascript/api/manifest/enabled) 元素设置为 `false`。 请参阅上面的示例。

第二步是分配处理程序。 通常在 **Office.onReady 函** 数中执行此操作，如以下示例所示，该示例将 (在后续步骤中创建的处理程序分配) 到工作表中所有图表的 **onActivated** 和 **onDeactivated** 事件。

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        const charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

第三步是定义 `enableChartFormat` 处理程序。 以下是一个简单示例，请参阅下面的[最佳做法：测试控件状态错误](#best-practice-test-for-control-status-errors)，以获取更改控件状态的更可靠方法。

```javascript
function enableChartFormat() {
    const button = {
                  id: "ChartFormatButton", 
                  enabled: true
                 };
    const parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    const parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    const ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

第四步是定义 `disableChartFormat` 处理程序。 除了将按钮对象的 **enabled** 属性设置为 `false` 之外，其他操作与 `enableChartFormat` 相同。

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>同时切换选项卡可见性和按钮的启用状态

**requestUpdate** 方法还用于切换自定义上下文选项卡的可见性。有关此代码和示例代码的详细信息，请参阅 [Office 加载项中的“创建自定义上下文”选项卡](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time)。

## <a name="best-practice-test-for-control-status-errors"></a>最佳做法：测试控件状态错误

在某些情况下，调用 `requestUpdate` 后，功能区不会重画，因此控件的可单击状态不会发生更改。 因此，加载项的最佳做法是跟踪其控件的状态。 加载项应符合以下规则。

1. 每当调用 `requestUpdate` 时，代码都应记录自定义按钮和菜单项的预期状态。
2. 单击自定义控件时，处理程序中的第一个代码应检查该按钮是否应为可单击按钮。 如果不是，则该代码应报告或记录错误，然后再次尝试将按钮设置为预期状态。

以下示例显示用于禁用按钮和记录按钮状态的函数。 请注意，`chartFormatButtonEnabled` 是全局布尔变量，其初始化为与清单中按钮的 [Enabled](/javascript/api/manifest/enabled) 元素相同的值。

```javascript
function disableChartFormat() {
    const button = {
                  id: "ChartFormatButton", 
                  enabled: false
                 };
    const parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    const parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    const ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

以下示例显示按钮的处理程序如何测试按钮的错误状态。 请注意，`reportError` 是用于显示或记录错误的函数。

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

## <a name="error-handling"></a>错误处理

在某些情况下，Office 无法更新功能区，并将返回错误。 例如，如果升级了加载项，并且升级后的加载项具有一组不同的自定义加载项命令，则必须关闭并重新打开 Office 应用程序。 在此之前，`requestUpdate` 方法将返回错误 `HostRestartNeeded`。 以下是如何处理此错误的示例。 在此示例中，`reportError` 方法向用户显示错误。

```javascript
function disableChartFormat() {
    try {
        const button = {
                      id: "ChartFormatButton", 
                      enabled: false
                     };
        const parentGroup = {
                           id: "MyGroup",
                           controls: [button]
                          };
        const parentTab = {
                         id: "CustomChartTab", 
                         groups: [parentGroup]
                        };
        const ribbonUpdater = {tabs: [parentTab]};
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

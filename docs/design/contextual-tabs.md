---
title: 在加载项中Office上下文选项卡
description: 了解如何将自定义上下文选项卡添加到Office外接程序。
ms.date: 09/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4d02d0c2201d4963be147c519a58c40cf3d5341c
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/18/2021
ms.locfileid: "59443550"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>在加载项中Office上下文选项卡

上下文选项卡是活动区中隐藏的选项卡控件Office，当文档中发生指定事件时，该选项卡Office显示。 例如 **，当选择** 表时，Excel功能区上出现的"表设计"选项卡。 在加载项中添加自定义Office，并创建更改可见性的事件处理程序，指定它们何时可见或隐藏。  (但是，自定义上下文选项卡不会响应焦点更改。) 

> [!NOTE]
> 本文假定你熟悉以下文档。 如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。
>
> - [加载项命令的基本概念](add-in-commands.md)

[!INCLUDE [Animation of contextual tabs and enabling buttons](../includes/animation-contextual-tabs-enable-button.md)]

> [!IMPORTANT]
> 自定义上下文选项卡当前仅在 Excel且仅在这些平台和内部版本上受支持。
>
> - Excel订阅Windows (Microsoft 365版本) 版本 2102 (内部版本 13801.20294) 或更高版本。
> - Excel 网页版

> [!NOTE]
> 自定义上下文选项卡仅适用于支持以下要求集的平台。 有关要求集以及如何使用它们，请参阅指定Office应用程序和 API[要求](../develop/specify-office-hosts-and-api-requirements.md)。
>
> - [RibbonApi 1.2](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> 可以在代码中使用运行时检查来测试用户的主机和平台组合是否支持这些要求集，如指定 Office[应用程序和 API 要求中所述](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)。  (在清单中指定要求集的技术（也如该文章中所述）当前对 RibbonApi 1.2.) 不起作用。或者，当不支持自定义上下文选项卡时，您可以实现备用 [UI 体验](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

## <a name="behavior-of-custom-contextual-tabs"></a>自定义上下文选项卡的行为

自定义上下文选项卡的用户体验遵循内置上下文选项卡Office模式。 以下是放置自定义上下文选项卡的基本原则。

- 当自定义上下文选项卡可见时，它将显示在功能区的右端。
- 如果同时显示外接程序中的一个或多个内置上下文选项卡和一个或多个自定义上下文选项卡，则自定义上下文选项卡始终位于所有内置上下文选项卡的右侧。
- 如果您的外接程序具有多个上下文选项卡，并且存在多个上下文，并且存在多个上下文可见，则它们按照在外接程序中定义的顺序显示。  (方向与语言的方向Office相同;也就是说，在从左到右语言中为从左到右，但从右到左使用从右到左的语言。) 请参阅定义选项卡上出现的组和控件[](#define-the-groups-and-controls-that-appear-on-the-tab)，详细了解如何定义它们。
- 如果多个加载项具有特定上下文中可见的上下文选项卡，则它们按加载项的启动顺序显示。
- 与 *自定义* 核心选项卡不同，自定义上下文选项卡不会Office应用程序的功能区。 它们仅存在于Office运行加载项的文档中。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>在外接程序中添加上下文选项卡的主要步骤

以下是在外接程序中添加自定义上下文选项卡的主要步骤。

1. 将外接程序配置为使用共享运行时。
1. 定义选项卡及其上出现的组和控件。
1. 使用"上下文"选项卡注册Office。
1. 指定选项卡可见时的情况。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>将外接程序配置为使用共享运行时

添加自定义上下文选项卡要求加载项使用共享运行时。 有关详细信息，请参阅 [将加载项配置为使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>定义显示在选项卡上的组和控件

与在清单中用 XML 定义的自定义核心选项卡不同，自定义上下文选项卡是在运行时使用 JSON blob 定义的。 代码将 blob 解析为 JavaScript 对象，然后将该对象传递给[Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)方法。 自定义上下文选项卡仅存在于当前运行加载项的文档中。 这不同于在安装外接程序时添加到 Office 应用程序功能区中的自定义核心选项卡，而当打开另一个文档时，这些选项卡将一直存在。 此外 `requestCreateControls` ，方法只能在加载项会话中运行一次。 如果再次调用它，将引发错误。

> [!NOTE]
> JSON blob 的属性和子属性 (和键名称) 的结构大致与清单 XML 中 [CustomTab](../reference/manifest/customtab.md) 元素及其后代元素的结构平行。

我们将分步构造上下文选项卡 JSON blob 的示例。 上下文选项卡 JSON 的完整架构位于 [dynamic-ribbon.schema.json 中](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 如果你使用 Visual Studio Code，可以使用此文件获取IntelliSense并验证 JSON。 有关详细信息，请参阅使用[JSON 架构和Visual Studio Code编辑 JSON。](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)

1. 首先，创建一个包含名为 和 的两个数组属性的 JSON `actions` 字符串 `tabs` 。 `actions`数组是上下文选项卡上的控件可以执行的所有函数的规范。数组 `tabs` 定义一个或多个上下文选项卡，*最多 20 个*。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. 这个上下文选项卡的简单示例将只有一个按钮，因此只有一个操作。 将以下内容添加为数组的唯一 `actions` 成员。 关于此标记，请注意：

    - `id`和 `type` 属性是必需的。
    - 的值 `type` 可以是"ExecuteFunction"或"ShowTaskpane"。
    - `functionName`只有当 的值为 时，才使用 `type` 属性 `ExecuteFunction` 。 它是 FunctionFile 中定义的函数的名称。 有关 FunctionFile 的信息，请参阅 [加载项命令的基本概念](add-in-commands.md)。
    - 在稍后的步骤中，您将此操作映射到上下文选项卡上的按钮。

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. 将以下内容添加为数组的唯一 `tabs` 成员。 关于此标记，请注意：

    - `id` 属性是必需的。 使用外接程序中所有上下文选项卡中唯一的简短描述性 ID。
    - `label` 属性是必需的。 它是一个用户友好字符串，用作上下文选项卡的标签。
    - `groups` 属性是必需的。 它定义将显示在选项卡上的控件组。它必须至少有一个成员且不超过 *20* 个。  (自定义上下文选项卡上可以具有的控件数量也有限制，这也会限制你拥有多少个组。 有关详细信息，请参阅下一步。) 

    > [!NOTE]
    > tab 对象还可以具有一个可选属性，该属性指定在外接程序启动时选项卡 `visible` 是否立即可见。 由于上下文选项卡通常是隐藏的，直到用户事件触发其可见性 (如用户在文档中选择某种类型的实体) 因此当不存在时，该属性默认为 `visible` `false` 。 在稍后的部分中，我们将展示如何设置 属性 `true` 以响应事件。

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. 在简单正在进行的示例中，上下文选项卡只有一个组。 将以下内容添加为数组的唯一 `groups` 成员。 关于此标记，请注意：

    - 所有属性都是必需的。
    - 属性在选项卡的所有组中必须是唯一的。 `id` 请使用简短的描述性 ID。
    - `label`是用作组标签的用户友好字符串。
    - 属性的值是一组对象，这些对象根据功能区的大小以及应用程序窗口的功能区Office `icon` 图标。
    - `controls`属性的值是指定组中按钮和菜单的对象数组。 必须至少有一个。

    > [!IMPORTANT]
    > *整个选项卡上的控件总数不能超过 20 个。* 例如，可以有 3 个组，每个组有 6 个控件，第四个组有 2 个控件，但不能有 4 个组，每个组有 6 个控件。  

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. 每个组必须具有至少两个大小的图标：32x32 像素和 80x80 像素。 （可选）还可以具有大小为 16x16 像素、20x20 像素、24x24 像素、40x40 像素、48x48 像素和 64x64 像素的图标。 Office根据功能区的大小和应用程序窗口的大小决定Office图标。 将以下对象添加到图标数组。  (如果窗口和功能区的大小足以使组中至少有一个控件显示，则不显示任何组图标。 例如，在缩小和展开 Word 窗口时，观察 Word 功能区上的 **"** 样式"组。) 关于此标记，请注意：

    - 这两个属性都是必需的。
    - `size`属性度量单位是像素。 图标始终为正方形，因此数字同时为高度和宽度。
    - `sourceLocation`属性指定图标的完整 URL。

    > [!IMPORTANT]
    > 与从开发环境移动到生产 (（如将域从 localhost 更改为 contoso.com) ）时，通常必须更改加载项清单中的 URL 一样，您还必须更改上下文选项卡 JSON 中的 URL。

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. 在我们的简单正在进行的示例中，该组只有一个按钮。 将以下对象添加为数组的唯一 `controls` 成员。 关于此标记，请注意：

    - 除 之外的所有属性 `enabled` 都是必需的。
    - `type` 指定控件的类型。 值可以是"Button"、"Menu"或"MobileButton"。
    - `id` 可以是最多 125 个字符。
    - `actionId` 必须是数组中定义的操作 `actions` ID。  (请参阅本节的步骤 1。) 
    - `label` 是用作按钮标签的用户友好字符串。
    - `superTip` 表示工具提示的丰富形式。 和 `title` `description` 属性都是必需的。
    - `icon` 指定按钮的图标。 前面有关组图标的备注也适用于此处。
    - `enabled` (可选) 指定在上下文选项卡启动时是否启用按钮。 如果不存在，则默认为 `true` 。

    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```

下面是 JSON blob 的完整示例。

```json
`{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>使用 requestCreateControls Office上下文选项卡

上下文选项卡通过调用[Office.ribbon.requestCreateControls 方法Office注册上下文](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)选项卡。 这通常在分配给 方法的函数中或在 方法 `Office.initialize` 中 `Office.onReady` 完成。 有关这些方法和初始化外接程序的更多信息，请参阅[Initialize your Office Add-in](../develop/initialize-add-in.md)。 但是，您可以在初始化后随时调用 该方法。

> [!IMPORTANT]
> `requestCreateControls`在加载项的给定会话中，只能调用方法一次。 如果再次调用，将引发错误。

示例如下。 请注意，必须先使用 方法将 JSON 字符串转换为 JavaScript 对象，然后才能 `JSON.parse` 将其传递给 JavaScript 函数。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>使用 requestUpdate 指定选项卡可见时上下文

通常，当用户启动的事件更改外接程序上下文时，应显示自定义上下文选项卡。 请考虑在激活工作簿的默认工作表上的图表时（并且仅在激活 (时，选项卡Excel可见) 的情况。

首先分配处理程序。 这通常在 方法中完成，如以下示例所示，该示例将 (在稍后步骤) 中创建的处理程序分配给工作表中所有图表的 和 `Office.onReady` `onActivated` `onDeactivated` 事件。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

接下来，定义处理程序。 下面是 的一个简单示例，但请参阅本文稍后介绍的处理 `showDataTab` [HostRestartNeeded](#handle-the-hostrestartneeded-error) 错误，了解函数的更可靠版本。 关于此代码，请注意以下几点：

- Office 控制何时更新功能区的状态。 the [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestUpdate_input_) method queues a request to update. 方法将在请求排入队列后（而不是功能区实际更新时） `Promise` 解析对象。
- 该方法的参数是 `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) 对象， (1) 按照 *JSON* 中指定的 ID 指定选项卡， (2) 指定选项卡的可见性。
- 如果你有多个自定义上下文选项卡应在同一上下文中可见，则只需向数组添加其他选项卡 `tabs` 对象。

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

隐藏选项卡的处理程序几乎完全相同，只是它将 `visible` 属性设置回 `false` 。

JavaScript Office还提供多个 (类型的) ，以便更轻松地构造 `RibbonUpdateData` 对象。 以下是 `showDataTab` TypeScript 中的 函数，它使用这些类型。

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>切换选项卡可见性和按钮的启用状态

方法还用于切换自定义上下文选项卡或自定义核心选项卡上自定义按钮的启用或 `requestUpdate` 禁用状态。有关详细信息，请参阅启用 [和禁用加载项命令](disable-add-in-commands.md)。 在某些情况下，你可能希望同时更改选项卡的可见性和按钮的启用状态。 为此，请调用 `requestUpdate` 。 下面是一个示例，其中核心选项卡上的按钮在显示上下文选项卡的同时启用。

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
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
            ]}
        ]
    });
}
```

在下面的示例中，启用的按钮位于可见上下文选项卡上。

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
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

## <a name="open-a-task-pane-from-contextual-tabs"></a>从上下文选项卡打开任务窗格

若要从自定义上下文选项卡上的按钮打开任务窗格，在 JSON 中创建具有 `type` 的 操作 `ShowTaskpane` 。 然后，定义一个 `actionId` 按钮，将 属性设置为 `id` 操作 。 这将打开清单中的 元素 `<Runtime>` 指定的默认任务窗格。

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

若要打开不是默认任务窗格的任何任务窗格，请指定操作 `sourceLocation` 定义中的属性。 在下面的示例中，从另一个按钮打开第二个任务窗格。

> [!IMPORTANT]
>
> - 如果 `sourceLocation` 为 操作指定了 ，则任务窗格 *不使用* 共享运行时。 它在新的 JavaScript 运行时中运行。
> - 共享运行时不能超过一个任务窗格，因此类型操作不能超过一个，因此 `ShowTaskpane` 不能忽略 `sourceLocation` 属性。

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    },
    {
      "id": "openTablesTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Tables",
      "supportPinning": false
      "sourceLocation": "https://MyDomain.com/myPage.html"
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            },
            {
                "type": "Button",
                "id": "CtxBt113",
                "actionId": "openTablesTaskpane",
                "enabled": false,
                "label": "Open Tables Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="localize-the-json-text"></a>本地化 JSON 文本

传递给 的 JSON blob 的本地化方式与自定义核心选项卡的清单标记本地化 (如从清单控件本地化 `requestCreateControls`) 。 [](../develop/localization.md#control-localization-from-the-manifest) 相反，本地化必须在运行时针对每个区域设置使用不同的 JSON blob。 建议您使用用于测试 `switch` [Office.context.displayLanguage 属性的](/javascript/api/office/office.context#displayLanguage)语句。 示例如下。

```javascript
function GetContextualTabsJsonSupportedLocale () {
    var displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

然后，代码调用 函数，获取传递给 的本地化 `requestCreateControls` blob，如以下示例所示。

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>自定义上下文选项卡的最佳实践

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>在不支持自定义上下文选项卡时实现备用 UI 体验

平台、应用程序Office和内部版本Office组合不支持 `requestCreateControls` 。 您的外接程序应设计为为在这些组合之一上运行外接程序的用户提供备用体验。 以下各节介绍提供回退体验的两种方法。

#### <a name="use-noncontextual-tabs-or-controls"></a>使用非上下文选项卡或控件

有一个清单元素 [OverriddenByRibbonApi，](../reference/manifest/overriddenbyribbonapi.md)旨在当外接程序在不支持自定义上下文选项卡的应用程序或平台上运行时，在外接程序中实现自定义上下文选项卡的回退体验。

使用此元素的最简单策略是，在清单中定义一个或多个自定义核心选项卡 (即，与外接程序中自定义上下文选项卡的功能区自定义项重复的非上下文自定义选项卡) 。 但添加为自定义核心选项卡上重复的 Group、Control 和 menu 元素的第 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` [](../reference/manifest/group.md)[](../reference/manifest/control.md) `<Item>` 一个子元素。 这样做的效果如下：

- 如果外接程序在支持自定义上下文选项卡的应用程序和平台上运行，则自定义核心组和控件将不会显示在功能区上。 相反，自定义上下文选项卡将在外接程序调用 方法时 `requestCreateControls` 创建。
- 如果外接程序在不支持 的应用程序或平台上运行，则元素会显示在自定义核心 `requestCreateControls` 选项卡上。

示例如下。 请注意，仅在不支持自定义上下文选项卡时，"MyButton"将显示在自定义核心选项卡上。 但是，无论自定义上下文选项卡是否受支持，都会显示父组和自定义核心选项卡。

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
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

有关更多示例，请参阅 [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)。

当父组或菜单标记为 时，它将不可见，并且当自定义上下文选项卡不受支持时，将忽略其 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 所有子标记。 因此，这些子元素中的任一元素是否具有 元素或其 `<OverriddenByRibbonApi>` 值是什么并不重要。 其含义是，如果菜单项或控件必须在所有上下文中可见，则不仅不应使用 标记它，而且其上级菜单和组也必须不按此方式 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` *标记*。

> [!IMPORTANT]
> 请勿使用 *标记* 组或菜单的所有子元素 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。 如果由于上一段给出的原因而标记父元素， `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 则这一点没有意义。 此外，如果在父选项卡上将 (或设置为) ，则无论自定义上下文选项卡是否受支持，父选项卡都会显示，但在支持自定义上下文选项卡时将为空 `<OverriddenByRibbonApi>` `false` 。 因此，如果支持自定义上下文选项卡时不应显示所有子元素，则使用 标记父元素，仅标记父元素 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>使用在指定的上下文中显示或隐藏任务窗格的 API

作为替代方法，加载项可以使用与自定义上下文选项卡上控件功能重复的 UI 控件定义 `<OverriddenByRibbonApi>` 任务窗格。然后，使用[Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__)和[Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__)方法来显示任务窗格，并且仅在支持上下文选项卡时显示该上下文选项卡。 有关如何使用这些方法的详细信息，请参阅显示或隐藏加载项Office[窗格](../develop/show-hide-add-in.md)。

### <a name="handle-the-hostrestartneeded-error"></a>处理 HostRestartNeeded 错误

在某些情况下，Office 无法更新功能区，并将返回错误。 例如，如果升级了加载项，并且升级后的加载项具有一组不同的自定义加载项命令，则必须关闭并重新打开 Office 应用程序。 在此之前，`requestUpdate` 方法将返回错误 `HostRestartNeeded`。 代码应处理此错误。 下面是操作方法的示例。 在此示例中，`reportError` 方法向用户显示错误。

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```

## <a name="resources"></a>资源

- [代码示例：在功能区上创建自定义上下文选项卡](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/office-contextual-tabs)
- Community上下文选项卡示例的演示

> [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]
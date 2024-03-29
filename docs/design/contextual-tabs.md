---
title: 在 Office 加载项中创建自定义上下文选项卡
description: 了解如何将自定义上下文选项卡添加到 Office 外接程序。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f43f6ec0a6ef3faef4c5e50d5da6d124124fe92
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810230"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>在 Office 加载项中创建自定义上下文选项卡

上下文选项卡是 Office 功能区中隐藏的选项卡控件，当 Office 文档中发生指定事件时，它将显示在选项卡行中。 例如，选择表格时 Excel 功能区上显示的“表格 **设计** ”选项卡。 在 Office 外接程序中包含自定义上下文选项卡，并通过创建更改可见性的事件处理程序来指定它们何时可见或隐藏。  (但是，自定义上下文选项卡不会响应焦点更改。) 

> [!NOTE]
> 本文假定你熟悉以下文档。 如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。
>
> - [加载项命令的基本概念](add-in-commands.md)

> [!IMPORTANT]
> 自定义上下文选项卡目前仅在 Excel 上受支持，并且仅在这些平台和版本中受支持。
>
> - Windows 上的 Excel：版本 2102 (内部版本 13801.20294) 或更高版本。
> - Mac 上的 Excel：版本 16.53.806.0 或更高版本。
> - Excel 网页版

> [!NOTE]
> 自定义上下文选项卡仅适用于支持以下要求集的平台。 有关要求集及其使用方式的详细信息，请参阅 [指定 Office 应用程序和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)。
>
> - [RibbonApi 1.2](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
> - [SharedRuntime 1.1](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)
>
> 可以使用代码中的运行时检查来测试用户的主机和平台组合是否支持这些要求集，如 [运行时检查方法和要求集支持](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)中所述。  (在清单中指定要求集的技术（本文中也进行了介绍）目前不适用于 RibbonApi 1.2.) 或者，当 [不支持自定义上下文选项卡时，可以实现备用 UI 体验](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

## <a name="behavior-of-custom-contextual-tabs"></a>自定义上下文选项卡的行为

自定义上下文选项卡的用户体验遵循内置 Office 上下文选项卡的模式。 下面是放置自定义上下文选项卡的基本原则。

- 当自定义上下文选项卡可见时，它将显示在功能区的右端。
- 如果同时显示一个或多个内置上下文选项卡和加载项中的一个或多个自定义上下文选项卡，则自定义上下文选项卡始终位于所有内置上下文选项卡的右侧。
- 如果外接程序具有多个上下文选项卡，并且存在多个可见的上下文，则它们按外接程序中的定义顺序显示。  (方向与 Office 语言的方向相同;也就是说， 在从左到右的语言中为从左到右，但在从右到左的语言中为从右到左。) 请参阅 [定义显示在选项卡上的组和控件](#define-the-groups-and-controls-that-appear-on-the-tab) ，了解有关如何定义组和控件的详细信息。
- 如果多个加载项具有在特定上下文中可见的上下文选项卡，则它们将按加载项的启动顺序显示。
- 自定义 *上下文* 选项卡与自定义核心选项卡不同，不会永久添加到 Office 应用程序的功能区。 它们仅存在于运行加载项的 Office 文档中。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>在加载项中包含上下文选项卡的主要步骤

下面是在外接程序中包含自定义上下文选项卡的主要步骤。

1. 将加载项配置为使用共享运行时。
1. 定义选项卡及其上显示的组和控件。
1. 将上下文选项卡注册到 Office。
1. 指定选项卡可见的情况。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>将加载项配置为使用共享运行时

添加自定义上下文选项卡需要外接程序使用 [共享运行时](../testing/runtimes.md#shared-runtime)。 有关详细信息，请参阅 [配置外接程序以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>定义显示在选项卡上的组和控件

与在清单中使用 XML 定义的自定义核心选项卡不同，自定义上下文选项卡是在运行时使用 JSON Blob 定义的。 代码将 Blob 分析为 JavaScript 对象，然后将该对象传递给 [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) 方法。 自定义上下文选项卡仅存在于当前运行加载项的文档中。 这与自定义核心选项卡不同，后者在安装加载项时添加到 Office 应用程序功能区，并在打开其他文档时保持存在。 此外， `requestCreateControls` 方法只能在加载项的会话中运行一次。 如果再次调用，则会引发错误。

> [!NOTE]
> JSON Blob 的属性和子属性的结构 (键名称) 大致与清单 XML 中 [CustomTab](/javascript/api/manifest/customtab) 元素及其后代元素的结构并行。

我们将逐步构造上下文选项卡 JSON Blob 的示例。 上下文选项卡 JSON 的完整架构位于 [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 如果使用的是 Visual Studio Code，则可以使用此文件获取 IntelliSense 并验证 JSON。 有关详细信息，请参阅[使用Visual Studio Code编辑 JSON - JSON 架构和设置](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)。

1. 首先创建包含两个名为 和 `tabs`的数组属性的 `actions` JSON 字符串。 数组 `actions` 是上下文选项卡上的控件可以执行的所有函数的规范。数组 `tabs` 定义一个或多个上下文选项卡， *最多 20 个*。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. 上下文选项卡的这个简单示例将只有一个按钮，因此只有一个操作。 添加以下内容作为数组的唯一 `actions` 成员。 关于此标记，请注意：

    - `id`和 `type` 属性是必需的。
    - 的值 `type` 可以是“ExecuteFunction”或“ShowTaskpane”。
    - `functionName`仅当 的`type``ExecuteFunction`值为 时，才使用 属性。 它是 FunctionFile 中定义的函数的名称。 有关 FunctionFile 的详细信息，请参阅 [外接程序命令的基本概念](add-in-commands.md)。
    - 在后面的步骤中，你将此操作映射到上下文选项卡上的按钮。

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
    ```

1. 添加以下内容作为数组的唯一 `tabs` 成员。 关于此标记，请注意：

    - `id` 属性是必需的。 使用在加载项中的所有上下文选项卡中唯一的简短描述性 ID。
    - `label` 属性是必需的。 它是一个用户友好的字符串，用作上下文选项卡的标签。
    - `groups` 属性是必需的。 它定义将在选项卡上显示的控件组。它必须至少有一个成员 *且不超过 20* 个。  (自定义上下文选项卡上的控件数量也有限制，并且也会限制你拥有的组数。 有关详细信息，请参阅下一步。) 

    > [!NOTE]
    > 选项卡对象还可以具有可选 `visible` 属性，该属性指定在加载项启动时是否立即显示选项卡。 由于上下文选项卡通常是隐藏的，直到用户事件触发其可见性 (例如用户在文档中选择某种类型的实体) ，因此当不存在时， `visible` 属性默认为 `false` 。 在后面的部分中，我们将演示如何将 属性 `true` 设置为 以响应事件。

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. 在简单的持续示例中，上下文选项卡只有一个组。 添加以下内容作为数组的唯一 `groups` 成员。 关于此标记，请注意：

    - 所有属性都是必需的。
    - 属性 `id` 在清单中的所有组中必须是唯一的。 使用最多 125 个字符的简短描述性 ID。
    - `label`是一个用户友好的字符串，用作组的标签。
    - 属性 `icon` 的值是对象数组，这些对象指定组将在功能区上具有的图标，具体取决于功能区和 Office 应用程序窗口的大小。
    - 属性 `controls` 的值是对象数组，这些对象指定组中的按钮和菜单。 必须至少有一个。

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

1. 每个组必须至少有两个大小（32x32 像素和 80x80 像素）的图标。 （可选）还可以具有大小为 16x16 像素、20x20 像素、24x24 像素、40x40 像素、48x48 像素和 64x64 像素的图标。 Office 根据功能区和 Office 应用程序窗口的大小决定要使用的图标。 将以下对象添加到图标数组。  (如果窗口和功能区大小足够大，以便至少显示组中的一个 *控件* ，则根本不显示组图标。 有关示例，请在缩小和展开 Word 窗口时观察 Word 功能区上的 **“样式”** 组。) 关于此标记，请注意：

    - 这两个属性都是必需的。
    - 度量 `size` 的属性单位为像素。 图标始终是方形的，因此数字是高度和宽度。
    - 属性 `sourceLocation` 指定图标的完整 URL。

    > [!IMPORTANT]
    > 正如从开发迁移到生产 (（例如将域从 localhost 更改为 contoso.com) ）时，通常必须更改加载项清单中的 URL 一样，还必须更改上下文选项卡 JSON 中的 URL。

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

1. 在我们的简单持续示例中，组只有一个按钮。 将以下 对象添加为数组的唯一 `controls` 成员。 关于此标记，请注意：

    - 除 之外 `enabled`的所有属性都是必需的。
    - `type` 指定控件的类型。 值可以是“Button”、“Menu”或“MobileButton”。
    - `id` 最多可以有 125 个字符。
    - `actionId` 必须是数组中定义的操作的 `actions` ID。  (请参阅本部分的步骤 1。) 
    - `label` 是一个用户友好的字符串，用作按钮的标签。
    - `superTip` 表示工具提示的丰富形式。 `title`和 `description` 属性都是必需的。
    - `icon` 指定按钮的图标。 前面有关组图标的注释也适用于此处。
    - `enabled` (可选) 指定在上下文选项卡启动时是否启用按钮。 如果不存在，则默认值为 `true`。

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

下面是 JSON Blob 的完整示例。

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>使用 requestCreateControls 向 Office 注册上下文选项卡

上下文选项卡通过调用 [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) 方法注册到 Office。 这通常在分配给 `Office.initialize` 的函数中完成，或使用 函数 `Office.onReady` 完成。 有关这些函数和初始化加载项的详细信息，请参阅 [初始化 Office 外接程序](../develop/initialize-add-in.md)。 但是，可以在初始化后随时调用 方法。

> [!IMPORTANT]
> 在 `requestCreateControls` 加载项的给定会话中，只能调用一次 方法。 如果再次调用，将引发错误。

示例如下。 请注意，必须先使用 方法将 JSON 字符串转换为 JavaScript 对象， `JSON.parse` 然后才能将其传递给 JavaScript 函数。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>使用 requestUpdate 指定选项卡何时可见的上下文

通常，当用户发起的事件更改加载项上下文时，应显示自定义上下文选项卡。 假设在 Excel 工作簿的默认工作表上激活图表 (时（且仅当）激活) 时，选项卡才可见。

首先分配处理程序。 这通常在 函数中 `Office.onReady` 完成，如以下示例中所示，该函数将后面的步骤 (创建的处理程序) 分配给 `onActivated` 工作表中所有图表的 和 `onDeactivated` 事件。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        const charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

接下来，定义处理程序。 下面是 的一个 `showDataTab`简单示例，但请参阅本文后面的 [处理 HostRestartNeeded 错误](#handle-the-hostrestartneeded-error) ，以获取更可靠的函数版本。 关于此代码，请注意以下几点：

- Office 控制何时更新功能区的状态。 [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) 方法将更新请求排队。 方法将在请求排队后立即解析 `Promise` 对象，而不是在功能区实际更新时解析对象。
- 方法的参数 `requestUpdate` 是 [一个 RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) 对象， (1 个) 按其 ID 完全 *按照 JSON 中指定的* ID 指定选项卡， (2) 指定选项卡的可见性。
- 如果有多个应在同一上下文中可见的自定义上下文选项卡，只需将其他选项卡对象添加到数组中 `tabs` 即可。

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

隐藏选项卡的处理程序几乎相同，只不过它将 属性设置`visible`回 。`false`

Office JavaScript 库还提供了多个接口， (类型) ，以便更轻松地构造`RibbonUpdateData` 对象。 下面是 `showDataTab` TypeScript 中的 函数，它利用这些类型。

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>同时切换选项卡可见性和按钮的启用状态

方法 `requestUpdate` 还用于切换自定义上下文选项卡或自定义核心选项卡上自定义按钮的启用或禁用状态。有关此内容的详细信息，请参阅 [启用和禁用外接程序命令](disable-add-in-commands.md)。 在某些情况下，你可能希望同时更改选项卡的可见性和按钮的启用状态。 为此，只需调用 `requestUpdate`一次 。 下面是一个示例，其中在启用核心选项卡上的按钮的同时，上下文选项卡可见。

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

在以下示例中，启用的按钮位于可见的上下文选项卡上。

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

若要从自定义上下文选项卡上的按钮打开任务窗格，请在 JSON 中创建包含 `type` 的 `ShowTaskpane`操作。 然后定义一个按钮，并将 `actionId` 属性设置为 `id` 操作的 。 这会打开由清单中的 元素指定的 **\<Runtime\>** 默认任务窗格。

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

若要打开任何不是默认任务窗格的任务窗格，请在操作的定义中指定属性 `sourceLocation` 。 在以下示例中，将从其他按钮打开第二个任务窗格。

> [!IMPORTANT]
>
> - `sourceLocation`为操作指定时，任务窗格 *不使用* 共享运行时。 它在新的单独运行时中运行。
> - 不能有一个任务窗格可以使用共享运行时，因此，类型为 `ShowTaskpane` 的多个操作不能省略 属性 `sourceLocation` 。

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

传递给 `requestCreateControls` 的 JSON Blob 的本地化方式与本地化自定义核心选项卡的清单标记的方式不同， (从 [清单) 控制本地化](../develop/localization.md#control-localization-from-the-manifest) 中所述。 相反，必须在运行时使用每个区域设置的不同 JSON Blob 进行本地化。 建议使用测试 `switch` [Office.context.displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) 属性的语句。 示例如下。

```javascript
function GetContextualTabsJsonSupportedLocale () {
    const displayLanguage = Office.context.displayLanguage;

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

然后，代码调用 函数以获取传递给 `requestCreateControls`的本地化 Blob，如以下示例所示。

```javascript
const contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>自定义上下文选项卡的最佳做法

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>在不支持自定义上下文选项卡时实现备用 UI 体验

平台、Office 应用程序和 Office 内部版本的某些组合不支持 `requestCreateControls`。 外接程序应设计为为在这些组合之一上运行加载项的用户提供替代体验。 以下部分介绍提供回退体验的两种方法。

#### <a name="use-noncontextual-tabs-or-controls"></a>使用非文本选项卡或控件

有一个清单元素 [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi)，它旨在创建外接程序中的回退体验，当外接程序在不支持自定义上下文选项卡的应用程序或平台上运行时，该外接程序实现自定义上下文选项卡。

使用此元素的最简单策略是定义自定义核心选项卡 (即清单中的 *非文本* 自定义选项卡) ，以复制外接程序中自定义上下文选项卡的功能区自定义。 但添加 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 为自定义核心选项卡上重复的 [组](/javascript/api/manifest/group)、 [控件](/javascript/api/manifest/control)和菜单 **\<Item\>** 元素的第一个子元素。 这样做的效果如下：

- 如果外接程序在支持自定义上下文选项卡的应用程序和平台上运行，则自定义核心组和控件不会出现在功能区上。 相反，当外接程序调用 `requestCreateControls` 方法时，将创建自定义上下文选项卡。
- 如果外接程序在 *不支持* `requestCreateControls`的应用程序或平台上运行，则元素会显示在自定义核心选项卡上。

示例如下。 请注意，仅当不支持自定义上下文选项卡时，“MyButton”才会显示在自定义核心选项卡上。 但是，无论是否支持自定义上下文选项卡，都会显示父组和自定义核心选项卡。

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
                <Control ... id="Contoso.MyButton1">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

有关更多示例，请参阅 [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi)。

当父组或菜单标记为 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`时，它不可见，并且当不支持自定义上下文选项卡时，将忽略其所有子标记。 因此，这些子元素中的任何一个是否具有 **\<OverriddenByRibbonApi\>** 元素或其值并不重要。 这意味着，如果菜单项或控件必须在所有上下文中可见，则不仅不应使用 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`标记，而且 *其上级菜单和组也不得以这种方式标记*。

> [!IMPORTANT]
> 不要使用 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`标记组或菜单 *的所有* 子元素。 如果父元素因上一段中给出的原因而标记为 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` ，则这是毫无意义的。 此外，如果省略 **\<OverriddenByRibbonApi\>** 父 (或将其设置为 `false`) ，则无论是否支持自定义上下文选项卡，都会显示父选项卡，但在支持自定义上下文选项卡时，这些选项卡将为空。 因此，如果支持自定义上下文选项卡时不应显示所有子元素，请使用 标记父级 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`。

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>使用在指定上下文中显示或隐藏任务窗格的 API

作为 的 **\<OverriddenByRibbonApi\>** 替代方法，加载项可以使用 UI 控件定义任务窗格，这些控件在自定义上下文选项卡上复制控件的功能。然后，使用 [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-showastaskpane-member(1)) 和 [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-hide-member(1)) 方法显示任务窗格（如果支持上下文选项卡）。 有关如何使用这些方法的详细信息，请参阅 [显示或隐藏 Office 外接程序的任务窗格](../develop/show-hide-add-in.md)。

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

- [代码示例：在功能区上创建自定义上下文选项卡](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs)
- 上下文选项卡示例的社区演示

> [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]

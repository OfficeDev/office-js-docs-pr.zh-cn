---
title: 在 Office 加载项中创建自定义上下文选项卡
description: 了解如何将自定义上下文选项卡添加到 Office 外接程序。
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: 12286ef675a938e4abd8dd3caa90cd97586cb6d7
ms.sourcegitcommit: 6a378d2a3679757c5014808ae9da8ababbfe8b16
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/15/2021
ms.locfileid: "49870635"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>在 Office 加载项中创建自定义上下文选项卡（预览）

上下文选项卡是 Office 功能区中的隐藏选项卡控件，在 Office 文档中发生指定事件时显示在选项卡行中。 例如 **，选择表** 时显示在 Excel 功能区上的"表设计"选项卡。 您可以通过创建更改可见性的事件处理程序，在 Office 外接程序中包括自定义上下文选项卡并指定它们何时可见或隐藏。  (，自定义上下文选项卡不会响应焦点更改。) 

> [!NOTE]
> 本文假定你熟悉以下文档。 如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。
>
> - [加载项命令的基本概念](add-in-commands.md)

> [!IMPORTANT]
> 自定义上下文选项卡为预览。 请在开发或测试环境中试验它们，但不要将其添加到生产外接程序。
>
> 自定义上下文选项卡当前仅在 Excel 上受支持，并且仅在以下平台和内部版本上受支持：
>
> - 仅适用于 Windows (Microsoft 365 上的 Excel，而不是永久许可证) ：版本 2011 (内部版本 13426.20274) 。 你的 Microsoft 365 订阅可能需要位于当前频道 [ (预览版) ](https://insider.office.com/join/windows) 以前称为"每月频道 (定向) "或"预览体验成员慢"。

> [!NOTE]
> 自定义上下文选项卡仅适用于支持以下要求集的平台。 有关要求集以及如何使用它们，请参阅"指定 Office 应用程序和[API 要求"。](../develop/specify-office-hosts-and-api-requirements.md)
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a>自定义上下文选项卡的行为

自定义上下文选项卡的用户体验遵循内置 Office 上下文选项卡的模式。 以下是放置自定义上下文选项卡的基础知识：

- 当自定义上下文选项卡可见时，它将显示在功能区的右端。
- 如果加载项中的一个或多个内置上下文选项卡和一个或多个自定义上下文选项卡同时可见，则自定义上下文选项卡始终位于所有内置上下文选项卡的右侧。
- 如果您的外接程序具有多个上下文选项卡，并且存在多个可见上下文，则它们按照在外接程序中定义的顺序显示。  (方向与 Office 语言的方向相同;也就是说，使用从左到右的语言从左到右，但从右到左使用从右到左的语言。) 请参阅"定义选项卡上出现的组和控件[](#define-the-groups-and-controls-that-appear-on-the-tab)"，详细了解如何定义它们。
- 如果多个加载项具有特定上下文中可见的上下文选项卡，则它们按加载项的启动顺序显示。
- 自定义 *上下文* 选项卡与自定义核心选项卡不同，不会永久添加到 Office 应用程序的功能区。 它们仅存在于运行加载项的 Office 文档中。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>在加载项中添加上下文选项卡的主要步骤

以下是在加载项中添加自定义上下文选项卡的主要步骤：

1. 将外接程序配置为使用共享运行时。
1. 定义选项卡及其上出现的组和控件。
1. 向 Office 注册上下文选项卡。
1. 指定选项卡可见时的情况。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>配置加载项以使用共享运行时

添加自定义上下文选项卡需要加载项使用共享运行时。 有关详细信息，请参阅 [配置外接程序以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>定义显示在选项卡上的组和控件

与使用清单中的 XML 定义的自定义核心选项卡不同，自定义上下文选项卡在运行时使用 JSON blob 定义。 代码将 blob 解析为 JavaScript 对象，然后将该对象传递给 [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) 方法。 自定义上下文选项卡仅存在于加载项当前运行的文档中。 这不同于在安装加载项时添加到 Office 应用程序功能区的自定义核心选项卡，当打开另一个文档时，这些选项卡仍保持显示状态。 此外 `requestCreateControls` ，此方法只能在加载项会话中运行一次。 如果再次调用它，将引发错误。

> [!NOTE]
> JSON blob 的属性和子属性 (和键名称) 的结构大致与清单 XML 中 [CustomTab](../reference/manifest/customtab.md) 元素及其后代元素的结构平行。

我们将分步构造上下文选项卡 JSON blob 的示例。  (上下文选项卡 JSON 的完整架构位于 [dynamic-ribbon.schema.js上](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 此链接在上下文选项卡的早期预览阶段可能无法运行。 如果链接不工作，您可以在 .) 上的草稿 [dynamic-ribbon.schema.js](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json)找到架构的最新草稿（如果您使用 Visual Studio Code，您可以使用此文件获取 IntelliSense 并验证 JSON。 有关详细信息，请参阅编辑 [JSON 和Visual Studio代码 - JSON 架构和设置](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)。


1. 首先创建一个 JSON 字符串，该字符串具有名为 和 的两个 `actions` 数组属性 `tabs` 。 该数组是上下文选项卡上的控件可以执行的所有函数 `actions` 的规范。数组 `tabs` 定义一个或多个上下文选项卡，最多 *10 个*。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. 上下文选项卡的这个简单示例将只有一个按钮，因此只有一个操作。 将以下内容添加为数组的唯一 `actions` 成员。 关于此标记，请注意：

    - 和 `id` `type` 属性是必需的。
    - 其值 `type` 可以是"ExecuteFunction"或"ShowTaskpane"。
    - 该属性 `functionName` 仅在值为 `type` `ExecuteFunction` 时使用。 它是 FunctionFile 中定义的函数的名称。 有关 FunctionFile 详细信息，请参阅 [外接程序命令的基本概念](add-in-commands.md)。
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
    - `groups` 属性是必需的。 它定义将在选项卡上出现的控件组。它必须至少有一个成员，且 *不超过 20 个*。  (自定义上下文选项卡上可以具有的控件数量也具有一些限制，这也会限制你拥有多少个组。 有关详细信息，请参阅下一步。) 

    > [!NOTE]
    > Tab 对象还可以具有一个可选属性，该属性指定在加载项启动时选项卡是否立即 `visible` 可见。 由于上下文选项卡通常是隐藏的，直到用户事件触发其可见性 (例如用户在文档中选择某种类型的实体) ，该属性默认为不存在时 `visible` `false` 。 在稍后的部分中，我们将展示如何设置该属性 `true` 以响应事件。

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. 在简单的正在进行的示例中，上下文选项卡只有一个组。 将以下内容添加为数组的唯一 `groups` 成员。 关于此标记，请注意：

    - 所有属性都是必需的。
    - 该属性在选项卡的所有组中必须是唯一的。 `id` 请使用简短的描述性 ID。
    - 这是 `label` 一个用户友好字符串，用作组的标签。
    - 该属性的值是一组对象，这些对象根据功能区的大小和 Office 应用程序窗口指定组将在功能区上具有 `icon` 的图标。
    - `controls`该属性的值是指定组中按钮和菜单的对象数组。 组中必须至少有一个且 *不超过 6 个*。

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

1. 每个组都必须具有至少两个大小的图标：32x32 像素和 80x80 像素。 （可选）还可以具有大小为 16x16 像素、20x20 像素、24x24 像素、40x40 像素、48x48 像素和 64x64 像素的图标。 Office 根据功能区的大小和 Office 应用程序窗口决定使用哪个图标。 将以下对象添加到图标数组。  (如果窗口和功能区大小足够大，组中至少有一个控件可以显示，则不显示任何组图标。 例如，在缩小和展开 Word 窗口时，观察 Word 功能区上的 **Styles** 组。) 关于此标记，请注意：

    - 这两个属性都是必需的。
    - `size`属性度量单位为像素。 图标始终为正方形，因此数字为高度和宽度。
    - 该属性 `sourceLocation` 指定图标的完整 URL。

    > [!IMPORTANT]
    > 与在从开发环境移动到生产 (（如将域从 localhost 更改为 contoso.com) ）时，通常必须更改加载项清单中的 URL 一样，您还必须更改上下文选项卡 JSON 中的 URL。

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

    - 除属性外 `enabled` ，其他所有属性都是必需的。
    - `type` 指定控件的类型。 值可以是"Button"、"Menu"或"MobileButton"。
    - `id` 最多为 125 个字符。 
    - `actionId` 必须是数组中定义的操作 `actions` ID。  (请参阅本节的步骤 1.) 
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
 
以下是 JSON blob 的完整示例：

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

上下文选项卡通过调用[Office.ribbon.requestCreateControls 方法注册到 Office。](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) 这通常在分配给方法的函数中完成， `Office.initialize` 或随方法 `Office.onReady` 一起完成。 有关这些方法和初始化外接程序的更多信息，请参阅["初始化 Office 外接程序"。](../develop/initialize-add-in.md) 但是，可以在初始化后随时调用该方法。

> [!IMPORTANT]
> `requestCreateControls`该方法只能在加载项的给定会话中调用一次。 如果再次调用错误，将引发错误。

示例如下。 请注意，必须先使用该方法将 JSON 字符串转换为 JavaScript 对象，然后才能将其 `JSON.parse` 传递给 JavaScript 函数。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>使用 requestUpdate 指定选项卡何时可见

通常，当用户启动的事件更改加载项上下文时，应显示自定义上下文选项卡。 考虑在激活 Excel 工作簿的默认工作表上的图表 (且仅在激活时，选项卡) 可见。

首先分配处理程序。 此方法中通常完成此操作，如以下示例所示，该示例将 (步骤) 中创建的处理程序分配给工作表中所有图表的和 `Office.onReady` `onActivated` `onDeactivated` 事件。

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

接下来，定义处理程序。 下面是一个简单示例，但请参阅本文稍后介绍的"处理 `showDataTab` [HostRestartNeeded](#handling-the-hostrestartneeded-error) 错误"，了解函数的更可靠版本。 关于此代码，请注意以下几点：

- Office 控制何时更新功能区的状态。 [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)方法将更新请求排成队列。 一旦将请求排入队列，该方法将解析该对象，而不是功能 `Promise` 区实际更新时。
- 该方法的参数是 `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) 对象， (1) 按其 ID 指定 *选项卡* ，而 (2) 指定选项卡的可见性。
- 如果多个自定义上下文选项卡应在同一上下文中可见，只需向数组添加其他 `tabs` 选项卡对象。

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

隐藏选项卡的处理程序几乎完全相同，只是将 `visible` 该属性设置回 `false` 。

Office JavaScript 库还提供了多个 (类型的) ，以便更轻松地构造 `RibbonUpdateData` 对象。 以下是 `showDataTab` TypeScript 中的函数，它使用这些类型。

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>同时切换选项卡可见性和按钮的启用状态

该方法还用于切换自定义上下文选项卡或自定义核心选项卡上自定义按钮的启用 `requestUpdate` 或禁用状态。有关此内容的详细信息，请参阅["启用和禁用外接程序命令"。](disable-add-in-commands.md) 在某些情况下，你可能希望同时更改选项卡的可见性和按钮的启用状态。 可以通过单个调用来此操作 `requestUpdate` 。 下面是一个示例，在使上下文选项卡可见的同时启用核心选项卡上的按钮。

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

在下面的示例中，启用的按钮位于要显示上下文选项卡的同一个选项卡上。

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

## <a name="localizing-the-json-blob"></a>本地化 JSON blob

传递给的 JSON blob 的本地化方式与自定义核心选项卡的清单标记的本地化方式不同 (如清单控件本地化中所述 `requestCreateControls`) 。 [](../develop/localization.md#control-localization-from-the-manifest) 相反，本地化必须在运行时针对每个区域设置使用不同的 JSON blob。 建议您使用用于测试 `switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) 属性的语句。 示例如下：

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

然后，代码调用该函数，获取传递给的本地化 `requestCreateControls` blob，如以下示例所示：

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="handling-the-hostrestartneeded-error"></a>处理 HostRestartNeeded 错误

在某些情况下，Office 无法更新功能区，并将返回错误。 例如，如果升级了加载项，并且升级后的加载项具有一组不同的自定义加载项命令，则必须关闭并重新打开 Office 应用程序。 在此之前，`requestUpdate` 方法将返回错误 `HostRestartNeeded`。 以下是如何处理此错误的示例。 在此示例中，`reportError` 方法向用户显示错误。

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
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

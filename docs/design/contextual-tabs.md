---
title: 在 Office 加载项中创建自定义上下文选项卡
description: 了解如何将自定义上下文选项卡添加到 Office 外接程序。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 0badd779f22edc9b4659908409764bea1cde44f5
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237719"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>在 Office 加载项中创建自定义上下文选项卡（预览）

上下文选项卡是 Office 功能区中的隐藏选项卡控件，当 Office 文档中发生指定事件时，该控件显示在选项卡行中。 例如 **，选择表** 时 Excel 功能区上出现的"表设计"选项卡。 您可以通过创建更改可见性的事件处理程序，在 Office 外接程序中包括自定义上下文选项卡并指定它们何时可见或隐藏。  (但是，自定义上下文选项卡不会响应焦点更改。) 

> [!NOTE]
> 本文假定你熟悉以下文档。 如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。
>
> - [加载项命令的基本概念](add-in-commands.md)

> [!IMPORTANT]
> 自定义上下文选项卡为预览版。 请在开发或测试环境中试验它们，但不要将其添加到生产外接程序。
>
> 自定义上下文选项卡当前仅在 Excel 上受支持，并且仅在以下平台和内部版本上受支持：
>
> - 仅适用于 Windows (Microsoft 365 上的 Excel，而不是永久许可证) ：版本 2011 (内部版本 13426.20274) 。 你的 Microsoft 365 订阅可能需要位于当前频道 [ (预览版) ](https://insider.office.com/join/windows) 以前称为"每月频道 (定向) "或"预览体验成员慢"。

> [!NOTE]
> 自定义上下文选项卡仅适用于支持以下要求集的平台。 有关要求集以及如何使用它们，请参阅"指定 Office 应用程序和[API 要求"。](../develop/specify-office-hosts-and-api-requirements.md)
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a>自定义上下文选项卡的行为

自定义上下文选项卡的用户体验遵循内置 Office 上下文选项卡的模式。 下面是放置自定义上下文选项卡的基础知识：

- 当自定义上下文选项卡可见时，它将显示在功能区的右端。
- 如果加载项中的一个或多个内置上下文选项卡和一个或多个自定义上下文选项卡同时可见，则自定义上下文选项卡始终位于所有内置上下文选项卡的右侧。
- 如果您的外接程序具有多个上下文选项卡，并且存在多个上下文，并且多个上下文可见，则它们按照在外接程序中定义的顺序显示。  (方向与 Office 语言的方向相同;也就是说，使用从左到右的语言从左到右，但从右到左使用从右向左的语言。) 请参阅"定义选项卡上出现的组和控件"，详细了解如何定义它们。 [](#define-the-groups-and-controls-that-appear-on-the-tab)
- 如果多个加载项具有特定上下文中可见的上下文选项卡，则它们按加载项启动的顺序显示。
- 自定义 *上下文* 选项卡与自定义核心选项卡不同，不会永久添加到 Office 应用程序的功能区。 它们仅存在于运行加载项的 Office 文档中。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>在加载项中添加上下文选项卡的主要步骤

以下是在加载项中添加自定义上下文选项卡的主要步骤：

1. 将外接程序配置为使用共享运行时。
1. 定义选项卡及其上出现的组和控件。
1. 向 Office 注册上下文选项卡。
1. 指定选项卡可见时的情况。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>配置外接程序以使用共享运行时

添加自定义上下文选项卡需要加载项使用共享运行时。 有关详细信息，请参阅 [配置外接程序以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>定义显示在选项卡上的组和控件

与使用清单中的 XML 定义的自定义核心选项卡不同，自定义上下文选项卡在运行时使用 JSON blob 定义。 代码将 blob 解析为 JavaScript 对象，然后将该对象传递给 [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) 方法。 自定义上下文选项卡仅存在于加载项当前运行的文档中。 这不同于在安装加载项时添加到 Office 应用程序功能区的自定义核心选项卡，当打开另一个文档时，这些选项卡仍保持显示状态。 此外，此方法只能在加载项的会话中运行 `requestCreateControls` 一次。 如果再次调用它，将引发错误。

> [!NOTE]
> JSON blob 的属性和子属性 (和键名称) 的结构与清单 XML 中 [CustomTab](../reference/manifest/customtab.md) 元素及其后代元素的结构大致平行。

我们将分步构造上下文选项卡 JSON blob 的示例。  (上下文选项卡 JSON 的完整架构位于 [dynamic-ribbon.schema.js上](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 对于上下文选项卡，此链接在预览阶段可能未运行。 如果链接不工作，您可以在 .) 上的草稿 [dynamic-ribbon.schema.js](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json)找到架构的最新草稿（如果您当前使用 Visual Studio Code，可以使用此文件获取 IntelliSense 并验证 JSON。 有关详细信息，请参阅编辑 [JSON 和 Visual Studio Code - JSON 架构和设置](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)。


1. 首先创建一个 JSON 字符串，该字符串具有名为 and 的两个 `actions` 数组属性 `tabs` 。 该数组是上下文选项卡上的控件可以执行的所有函数 `actions` 的规范。数组 `tabs` 定义一个或多个上下文选项卡，最多 *20 个*。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. 此上下文选项卡的简单示例将只有一个按钮，因此只有一个操作。 将以下内容添加为数组的唯一 `actions` 成员。 关于此标记，请注意：

    - 和 `id` `type` 属性是必需的。
    - 其值 `type` 可以是"ExecuteFunction"或"ShowTaskpane"。
    - 该属性 `functionName` 仅在值为 `type` `ExecuteFunction` 时使用。 它是 FunctionFile 中定义的函数的名称。 有关 FunctionFile 的信息，请参阅 [外接程序命令的基本概念](add-in-commands.md)。
    - 在稍后的步骤中，将此操作映射到上下文选项卡上的按钮。

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
    - `groups` 属性是必需的。 它定义将在选项卡上出现的控件组。它必须至少有一个成员且不超过 *20 个*。  (自定义上下文选项卡上可以具有的控件数量也有限制，这也会限制您拥有多少个组。 有关详细信息，请参阅下一步。) 

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

1. 在简单的持续示例中，上下文选项卡只有一个组。 将以下内容添加为数组的唯一 `groups` 成员。 关于此标记，请注意：

    - 所有属性都是必需的。
    - 该属性在选项卡上的所有组中必须是唯一的。 `id` 请使用简短的描述性 ID。
    - 它是 `label` 一个用户友好字符串，用作组的标签。
    - 该属性的值是一组对象，这些对象根据功能区的大小和 Office 应用程序窗口指定该组将在功能区上具有 `icon` 的图标。
    - `controls`该属性的值是指定组中按钮和菜单的对象数组。 必须至少有一个。

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

1. 每个组都必须具有至少两个大小的图标：32x32 像素和 80x80 像素。 还可以具有大小为 16x16 像素、20x20 像素、24x24 像素、40x40 像素、48x48 像素和 64x64 像素的图标。 Office 根据功能区的大小和 Office 应用程序窗口决定使用哪个图标。 将以下对象添加到图标数组。  (如果窗口和功能区大小足以显示组上的至少一个控件，则不显示任何组图标。 例如，在缩小和展开 Word 窗口时，观察 Word 功能区上的 **Styles** 组。) 关于此标记，请注意：

    - 这两个属性都是必需的。
    - `size`属性度量单位是像素。 图标始终为正方形，因此数字为高度和宽度。
    - 该属性 `sourceLocation` 指定图标的完整 URL。

    > [!IMPORTANT]
    > 就像在从开发移动到生产 (（如将域从 localhost 更改为 contoso.com) ）时，通常必须更改外接程序清单中的 URL 一样，您还必须更改上下文选项卡 JSON 中的 URL。

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

    - 除属性外 `enabled` ，所有属性都是必需的。
    - `type` 指定控件的类型。 值可以是"Button"、"Menu"或"MobileButton"。
    - `id` 最多为 125 个字符。 
    - `actionId` 必须是数组中定义的操作 `actions` ID。  (请参阅本节的步骤 1.) 
    - `label` 是用作按钮标签的用户友好字符串。
    - `superTip` 表示工具提示的丰富形式。 和 `title` `description` 属性都是必需的。
    - `icon` 指定按钮的图标。 前面有关组图标的备注也适用于此处。
    - `enabled` (可选) 指定在上下文选项卡启动时是否启用按钮。 默认值（如果不存在）为 `true` 。 

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

上下文选项卡通过调用 [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) 方法注册到 Office。 这通常在分配给方法的函数中或通过该方法 `Office.initialize` `Office.onReady` 完成。 有关这些方法和初始化外接程序的更多信息，请参阅["初始化 Office 外接程序"。](../develop/initialize-add-in.md) 但是，可以在初始化后随时调用该方法。

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

通常，当用户启动的事件更改加载项上下文时，应显示自定义上下文选项卡。 请考虑一种方案，在激活 Excel 工作簿的默认工作表上的图表 (且仅在激活该选项卡时) 显示。

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

接下来，定义处理程序。 下面是一个简单的示例，请参阅本文稍后介绍的"处理 `showDataTab` [HostRestartNeeded](#handle-the-hostrestartneeded-error) 错误"，了解函数的更可靠版本。 关于此代码，请注意以下几点：

- Office 控制何时更新功能区的状态。 [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)方法将更新请求排成队列。 一旦将请求排入队列，该方法将解析该对象，而不是功能 `Promise` 区实际更新时。
- 该方法的参数是 `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) 对象，该对象 (1) 按照 *JSON* 中指定的 ID 指定选项卡， (2) 指定选项卡的可见性。
- 如果多个自定义上下文选项卡应在同一上下文中可见，只需向数组添加其他选项卡 `tabs` 对象。

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

Office JavaScript 库还提供多个 (类型) ，以便更轻松地构造 `RibbonUpdateData` 对象。 下面是 `showDataTab` TypeScript 中的函数，它使用这些类型。

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>同时切换选项卡可见性和按钮的启用状态

该方法还用于在自定义上下文选项卡或自定义核心选项卡上切换自定义按钮的启用或 `requestUpdate` 禁用状态。有关此内容的详细信息，请参阅["启用和禁用外接程序命令"。](disable-add-in-commands.md) 在某些情况下，你可能希望同时更改选项卡的可见性和按钮的启用状态。 可以通过单个调用来执行此操作 `requestUpdate` 。 下面是一个示例，在使上下文选项卡可见的同时启用核心选项卡上的按钮。

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

传递给的 JSON blob 的本地化方式与自定义核心选项卡的清单标记的本地化方式不同 (如清单中的控件本地化中所述 `requestCreateControls`) 。 [](../develop/localization.md#control-localization-from-the-manifest) 相反，本地化必须在运行时针对每个区域设置使用不同的 JSON blob。 建议您使用用于测试 `switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) 属性的语句。 示例如下：

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

然后，代码调用函数，获取传递给的本地化 `requestCreateControls` blob，如以下示例所示：

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>自定义上下文选项卡的最佳实践

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>在不支持自定义上下文选项卡时实现备用 UI 体验

平台、Office 应用程序和 Office 内部版本的组合不支持 `requestCreateControls` 。 您的外接程序应设计为为在这些组合之一上运行外接程序的用户提供备用体验。 以下各节介绍提供回退体验的两种方法。

#### <a name="use-noncontextual-tabs-or-controls"></a>使用非上下文选项卡或控件

有一个清单元素 [OverriddenByRibbonApi，](../reference/manifest/overriddenbyribbonapi.md)设计用于在外接程序中创建回退体验，该体验在外接程序在不支持自定义上下文选项卡的应用程序或平台上运行时实现自定义上下文选项卡。 

使用此元素的最简单策略是，在清单中定义一个或多个自定义核心选项卡 (，即，复制外接程序中自定义上下文选项卡的功能区自定义的非上下文自定义选项卡) 。 但添加 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 为 [CustomTab](../reference/manifest/customtab.md)的第一个子元素。 这样做的效果如下：

- 如果外接程序在支持自定义上下文选项卡的应用程序和平台上运行，则自定义核心选项卡将不会显示在功能区上。 相反，当加载项调用该方法时，将创建自定义上下文 `requestCreateControls` 选项卡。
- 如果加载项在不支持的应用程序或平台上运行，则自定义核心选项卡 `requestCreateControls` 会显示在功能区上。

下面是此简单策略的一个示例。

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
              <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
...
</OfficeApp>
```

此简单策略使用自定义核心选项卡，该选项卡将自定义上下文选项卡与它的子组和控件进行镜像，但可以使用更复杂的策略。 还可以 `<OverriddenByRibbonApi>` 将元素[](../reference/manifest/control.md)作为[](../reference/manifest/control.md#menu-dropdown-button-controls)第一 () 子元素添加到[组](../reference/manifest/group.md)和控件元素 (按钮类型和菜单) 元素[](../reference/manifest/control.md#button-control) `<Item>` 。 这一事实使您能够在各种自定义核心选项卡的各种组、按钮和菜单之间分发其他将显示在上下文选项卡上的组和控件。 示例如下。 请注意，仅在不支持自定义上下文选项卡时，"MyButton"才出现在自定义核心选项卡上。 但是，无论是否支持自定义上下文选项卡，都将显示父组和自定义核心选项卡。

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

当使用父选项卡、组或菜单进行标记时，它将不可见，并且当不支持自定义上下文选项卡时，将忽略它的所有子 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 标记。 因此，其中任何子元素是否具有元素 `<OverriddenByRibbonApi>` 或其值并不重要。 这意味着，如果菜单项、控件或组必须在所有上下文中可见，则不仅不应使用它进行标记，而且其上级菜单、组和选项卡也必须不这样 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` *标记*。

> [!IMPORTANT]
> 请勿使用 标记 *选项卡* 、组或菜单的所有子元素 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。 如果由于上一段落中给出的原因标记父元素， `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 则这一点没有意义。 此外，如果退出父 (或设置为 `<OverriddenByRibbonApi>`) ，则无论是否支持自定义上下文选项卡，父选项卡都会显示，但在支持自定义上下文选项卡时将为空。 `false` 因此，如果支持自定义上下文选项卡时不应显示所有子元素，则使用 标记父元素，并且仅标记父元素 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>使用在指定的上下文中显示或隐藏任务窗格的 API

作为替代方法，加载项可以使用与自定义上下文选项卡上的控件功能重复的 UI 控件定义 `<OverriddenByRibbonApi>` 任务窗格。然后，使用 [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) 和 [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) 方法在上下文选项卡受支持时（且仅在何时）显示任务窗格。 有关如何使用这些方法的详细信息，请参阅显示 [或隐藏 Office 加载项的任务窗格](../develop/show-hide-add-in.md)。

### <a name="handle-the-hostrestartneeded-error"></a>处理 HostRestartNeeded 错误

在某些情况下，Office 无法更新功能区，并将返回错误。 例如，如果升级了加载项，并且升级后的加载项具有一组不同的自定义加载项命令，则必须关闭并重新打开 Office 应用程序。 在此之前，`requestUpdate` 方法将返回错误 `HostRestartNeeded`。 代码应处理此错误。 下面是操作方法的一个示例。 在此示例中，`reportError` 方法向用户显示错误。

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

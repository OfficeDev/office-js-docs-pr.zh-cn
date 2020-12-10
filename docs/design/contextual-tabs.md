---
title: 在 Office 外接程序中创建自定义上下文选项卡
description: 了解如何将自定义上下文选项卡添加到 Office 外接程序。
ms.date: 11/20/2020
localization_priority: Normal
ms.openlocfilehash: d8617c7dd8748d15393c0e38c527062e5894e791
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/09/2020
ms.locfileid: "49612734"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>在 Office 外接程序中创建自定义上下文选项卡 (预览) 

上下文选项卡是 Office 功能区中的一个隐藏的选项卡控件，当 Office 文档中发生指定事件时，该选项卡将显示在 "选项卡" 行中。 例如，选择表时 Excel 功能区上显示的 " **表设计** " 选项卡。 您可以在 Office 外接程序中包含自定义上下文选项卡，并通过创建更改可见性的事件处理程序来指定它们何时可见或隐藏。  (但是，自定义上下文选项卡对焦点更改没有响应。 ) 

> [!NOTE]
> 本文假定你熟悉以下文档。 如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。
>
> - [加载项命令的基本概念](add-in-commands.md)

> [!IMPORTANT]
> "自定义上下文选项卡" 处于预览阶段。 请在开发或测试环境中试用它们，但不要将其添加到生产外接加载项中。
>
> 自定义上下文选项卡目前仅在 Excel 中受支持，并且仅在这些平台和生成上受支持：
>
> - Windows (的 Excel 仅适用于 Microsoft 365，而不是永久许可证) ：版本 2011 (内部版本 13426.20274) 。 您的 Microsoft 365 订阅可能需要在 [当前频道 (预览) ](https://insider.office.com/join/windows) 以前称为 "每月频道 (目标) " 或 "内幕慢速"。

> [!NOTE]
> 自定义上下文选项卡仅适用于支持以下要求集的平台。 有关要求集以及如何使用它们的详细信息，请参阅 [指定 Office 应用程序和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)。
>
> - [SharedRuntime 1。1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a>自定义上下文选项卡的行为

自定义上下文选项卡的用户体验遵循内置 Office 上下文选项卡的模式。 以下是放置自定义上下文选项卡的基本原则：

- 当自定义上下文选项卡可见时，它将显示在功能区的右端。
- 如果一个或多个内置上下文选项卡以及外接程序中的一个或多个自定义上下文选项卡同时可见，则自定义上下文选项卡将始终位于所有内置上下文选项卡的右侧。
- 如果你的外接程序有多个上下文选项卡，并且存在多个上下文选项卡，则它们将按其在外接程序中的定义顺序显示。  (方向与 Office 语言的方向相同，则为 [ ](#define-the-groups-and-controls-that-appear-on-the-tab) ; 否则为  。也就是说，从左到右的语言按从左到右的语言，但从右到左的语言为从右到左。 ) 请参阅定义显示在选项卡上的组和控件，了解有关如何定义它们的详细信息。
- 如果有多个加载项具有在特定上下文中可见的上下文选项卡，则它们将按加载项启动的顺序显示。
- 与自定义核心选项卡不同，自定义 *上下文* 选项卡不会永久添加到 Office 应用程序的功能区。 它们仅存在于运行外接程序的 Office 文档中。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>在外接程序中包含上下文选项卡的主要步骤

以下是在外接程序中包含自定义上下文选项卡的主要步骤：

1. 将加载项配置为使用共享运行时。
1. 定义选项卡以及显示在其上的组和控件。
1. 使用 Office 注册上下文选项卡。
1. 指定选项卡将可见的情况。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>将加载项配置为使用共享运行时

添加自定义上下文选项卡需要您的外接程序使用共享运行时。 有关详细信息，请参阅 [Configure a 外接程序以使用共享运行时](../excel/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>定义选项卡上显示的组和控件

与使用清单中的 XML 定义的自定义核心选项卡不同，自定义上下文选项卡是在运行时使用 JSON blob 定义的。 您的代码将 blob 解析为 JavaScript 对象，然后将该对象传递给 [requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) 方法。 自定义上下文选项卡仅存在于您的外接程序当前运行的文档中。 这不同于自定义核心选项卡，在安装加载项时，这些选项卡会添加到 Office 应用程序功能区中，并且在打开另一个文档时仍然存在。 此外，该 `requestCreateControls` 方法在外接程序的会话中只能运行一次。 如果再次调用该方法，则会引发错误。

> [!NOTE]
> JSON blob 的 properties 和子属性的结构 (和键名称) 大致与清单 XML 中的 [CustomTab](../reference/manifest/customtab.md) 元素及其后代元素的结构平行。

我们将构造一个上下文选项卡 JSON blob 的示例。  (上下文选项卡 JSON 的完整架构位于 [dynamic-ribbon.schema.js的](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 此链接可能无法在上下文选项卡的早期预览周期中工作。 如果链接未正常运行，则可以在 [草稿 dynamic-ribbon.schema.js上](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json)找到架构的最新草案。 ) 如果您在 Visual Studio Code 中工作，则可以使用此文件获取 IntelliSense 并验证您的 JSON。 有关详细信息，请参阅 [使用 Visual Studio CODE JSON 架构和设置编辑 JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)。


1. 首先，创建一个具有两个名为和的数组属性的 JSON 字符串 `actions` `tabs` 。 `actions`数组是上下文选项卡上的控件可以执行的所有函数的规范。`tabs`数组定义了一个或多个上下文选项卡，*最多可达 10* 个。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. 这一简单的上下文选项卡示例将仅包含一个按钮，因此仅有一个操作。 将以下项添加为数组中的唯一成员 `actions` 。 有关此标记的信息，请注意：

    - `id`和 `type` 属性是必需的。
    - 的值 `type` 可以是 "ExecuteFunction" 或 "ShowTaskpane"。
    - `functionName`仅当的值为时，才使用 `type` 属性 `ExecuteFunction` 。 它是在 FunctionFile 中定义的函数的名称。 有关 FunctionFile 的详细信息，请参阅 [外接程序命令的基本概念](add-in-commands.md)。
    - 在后续步骤中，将此操作映射到 "上下文" 选项卡上的一个按钮。

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. 将以下项添加为数组中的唯一成员 `tabs` 。 有关此标记的信息，请注意：

    - `id` 属性是必需的。 使用外接程序中的所有上下文选项卡中唯一的简短描述性 ID。
    - `label` 属性是必需的。 它是一个用户友好的字符串，用作上下文选项卡的标签。
    - `groups` 属性是必需的。 它定义将显示在选项卡上的控件组。它必须至少有一个成员 *，且不能超过20个*。  (还限制了在自定义上下文选项卡上可以拥有的控件数，同时也会限制您拥有的组数。 有关详细信息，请参阅下一步。 ) 

    > [!NOTE]
    > Tab 对象还可以具有一个可选 `visible` 属性，该属性指定在加载项启动时选项卡是否立即可见。 由于上下文选项卡通常是隐藏的，直到用户事件触发其可见性 (例如，用户在文档中选择某种类型的实体) ，而该 `visible` 属性默认 `false` 情况下不显示时。 在后面的部分中，我们将演示如何将属性设置为，以 `true` 响应事件。

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. 在简单的后续示例中，上下文选项卡仅有一个组。 将以下项添加为数组中的唯一成员 `groups` 。 有关此标记的信息，请注意：

    - 所有属性都是必需的。
    - 该 `id` 属性在选项卡中的所有组中必须是唯一的。使用简短的描述性 ID。
    - `label`是用户友好的字符串，用作组的标签。
    - 该 `icon` 属性的值是对象的数组，这些对象指定根据功能区和 Office 应用程序窗口的大小，组将在功能区上所具有的图标。
    - 该 `controls` 属性的值是指定组中的按钮和菜单的对象的数组。 组中必须至少有一个和 *不超过6个*。

    > [!IMPORTANT]
    > *"整个" 选项卡上的总控件数不能超过20。* 例如，可以有3个组，每个组具有6个控件，第四组具有2个控件，但您不能有4个组，每个组都有6个控件。  

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

1. 每个组都必须有至少两个大小的图标： 32x32 px 和 80x80 px。 （可选）还可以具有 16x16 px、20x20 px、24x24 px、40x40 px、48x48 px 和 64x64 px 大小的图标。 Office 根据功能区和 Office 应用程序窗口的大小决定要使用哪个图标。 将以下对象添加到图标数组中。  (如果窗口和功能区大小足以满足组中的至少一个 *控件* 的显示，则不会显示任何组图标。 有关示例，请查看 Word 功能区上的 **样式** 组，将其缩小并展开 word 窗口。有关此标记的 ) ，请注意：

    - 这两个属性都是必需的。
    - 该 `size` 属性的度量单位为像素。 图标始终为方形，因此该数字同时为高度和宽度。
    - `sourceLocation`属性指定图标的完整 URL。

    > [!IMPORTANT]
    > 在迁移到生产 (时，通常必须更改加载项清单中的 Url，如将域从 localhost 更改为 contoso.com) 中，则还必须更改上下文选项卡 JSON 中的 Url。

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

1. 在我们简单的示例中，组仅有一个按钮。 将以下对象添加为数组的唯一成员 `controls` 。 有关此标记的信息，请注意：

    - 除之外的所有属性 `enabled` 都是必需的。
    - `type` 指定控件的类型。 这些值可以是 "Button"、"Menu" 或 "MobileButton"。
    - `id` 最大可以为125个字符。 
    - `actionId` 必须是在数组中定义的操作的 ID `actions` 。  (请参阅本部分的步骤1。 ) 
    - `label` 是用户友好的字符串，用作按钮的标签。
    - `superTip` 代表一种丰富的工具提示形式。 `title`和 `description` 属性都是必需的。
    - `icon` 指定按钮的图标。 上面关于组图标的备注也适用于此处。
    - `enabled` (可选) 指定在启动上下文选项卡时是否启用按钮。 如果不存在，则为默认值 `true` 。 

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
 
下面是 JSON blob 的完整示例：

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
      "label": "Data",
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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>使用 requestCreateControls 注册带有 Office 的上下文选项卡

上下文选项卡通过调用 [requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) 方法在 Office 中注册。 这通常是在分配给或方法的函数中完成的 `Office.initialize` `Office.onReady` 。 有关这些方法和初始化外接程序的详细信息，请参阅 [初始化 Office 外接程序](../develop/initialize-add-in.md)。 不过，您可以在初始化后随时调用方法。

> [!IMPORTANT]
> `requestCreateControls`在外接程序的给定会话中，只能调用一次方法。 如果再次调用，则会引发错误。

示例如下。 请注意，必须使用方法将 JSON 字符串转换为 JavaScript 对象， `JSON.parse` 然后才能将其传递给 javascript 函数。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>在选项卡将与 requestUpdate 一起显示时指定上下文

通常情况下，当用户启动的事件更改加载项上下文时，将显示自定义上下文选项卡。 考虑在激活 Excel 工作簿) 的默认工作表上的图表 (时，选项卡应可见的情况。

首先分配处理程序。 这通常是在方法中完成的 `Office.onReady` ，如以下示例所示，在后续步骤中 (创建的处理程序分配) 到 `onActivated` `onDeactivated` 工作表中的所有图表的和事件。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string.
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

接下来，定义处理程序。 下面是一个简单的示例 `showDataTab` ，但请参阅本文稍后的 [处理 HostRestartNeeded 错误](#handling-the-hostrestartneeded-error) ，以获取更强健的函数版本。 关于此代码，请注意以下几点：

- Office 控制何时更新功能区的状态。 [RequestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)方法对要更新的请求进行排队。 该方法将在 `Promise` 对象排队请求（而不是功能区实际更新）时立即解析该对象。
- 方法的参数 `requestUpdate` 是一个[RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata)对象，该对象 (1) 按它的 ID 指定选项卡的 (ID。) 指定选项卡的可见性。
- 如果有多个自定义上下文选项卡应在相同上下文中可见，则只需向该数组中添加其他选项卡对象 `tabs` 。

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

隐藏选项卡的处理程序几乎完全相同，不同之处在于它将 `visible` 属性重新设置为 `false` 。

Office JavaScript 库还提供了多个 (类型) 接口，以便更轻松地构造 `RibbonUpdateData` 对象。 以下是 `showDataTab` TypeScript 中的函数，它使用这些类型。

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>同时切换选项卡可见性和按钮的启用状态

此 `requestUpdate` 方法还用于切换自定义上下文选项卡或自定义 "核心" 选项卡上的自定义按钮的启用或禁用状态。有关此内容的详细信息，请参阅 [Enable And Disable 外接程序命令](disable-add-in-commands.md)。 在某些情况下，您可能希望同时更改选项卡的可见性和按钮的启用状态。 您可以通过一次调用来执行此操作 `requestUpdate` 。 下面的示例演示了启用 "核心" 选项卡上的按钮时，将显示上下文选项卡。

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
                controls: [
                {
                    id: "MyButton",
                    enabled: true
                }
            ]}
        ]});
}
```

在下面的示例中，启用的按钮在相同的上下文选项卡上，使其可见。

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                controls: [
                    {
                        id: "MyButton",
                        enabled: true
                    }
                ]
            }
        ]});
}
```

## <a name="localizing-the-json-blob"></a>本地化 JSON blob

传递到的 JSON blob 的 `requestCreateControls` 本地化方式与自定义核心选项卡的清单标记的本地化方式相同 ([从清单) 的控件本地化中](../develop/localization.md#control-localization-from-the-manifest) 进行了说明。 相反，本地化必须在运行时对每个区域设置使用不同的 JSON blob。 建议使用对 `switch` [displayLanguage](/javascript/api/office/office.context#displayLanguage) 属性进行测试的语句。 示例如下：

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
                          "label": "Data",
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
                          "label": "Données",
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

然后，代码调用函数以获取传递给的本地化 blob `requestCreateControls` ，如下面的示例所示：

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

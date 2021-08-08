---
title: 加载项中的Office快捷方式
description: 了解如何将自定义键盘快捷方式（也称为组合键）Office加载项。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 0e622820dbf1c582188d038fc532236154b7ec7389a8e14b0ce0e982a6f67a8a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57081747"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a>将自定义键盘快捷方式添加到Office加载项

键盘快捷方式（也称为组合键）使加载项的用户能够更高效地工作。 键盘快捷方式通过提供鼠标的替代方法，还可以为残障人士改进加载项的辅助功能。

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> 若要从已启用键盘快捷方式的加载项的工作版本开始，请克隆并运行键盘快捷方式[Excel示例](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。 准备好向自己的加载项添加键盘快捷方式后，请继续阅读本文。

向加载项添加键盘快捷方式有三个步骤。

1. [配置加载项的清单](#configure-the-manifest)。
1. [创建或编辑快捷方式 JSON 文件](#create-or-edit-the-shortcuts-json-file) 以定义操作及其键盘快捷方式。
1. [添加](#create-a-mapping-of-actions-to-their-functions) [Office.actions.associate](/javascript/api/office/office.actions#associate) API 的一个或多个运行时调用，以将函数映射到每个操作。

## <a name="configure-the-manifest"></a>配置清单

清单有两个小更改需要进行。 一种是允许加载项使用共享运行时，另一种是指向定义键盘快捷方式的 JSON 格式文件。

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>将外接程序配置为使用共享运行时

添加自定义键盘快捷方式要求加载项使用共享运行时。 有关详细信息，请 [配置外接程序以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

### <a name="link-the-mapping-file-to-the-manifest"></a>将映射文件链接到清单

在 *紧* (不在) 元素内，添加 `<VersionOverrides>` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素。 将 `Url` 属性设置为项目中将在稍后步骤创建的 JSON 文件的完整 URL。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>创建或编辑快捷方式 JSON 文件

在项目中创建 JSON 文件。 确保文件的路径与为 `Url` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素的 属性指定的位置相匹配。 此文件将描述键盘快捷方式以及这些快捷方式将调用的操作。

1. 在 JSON 文件中，有两个数组。 actions 数组将包含定义要调用的操作的对象，快捷方式数组将包含将键组合映射到操作的对象。 下面是一个示例。
    ```json
    {
        "actions": [
            {
                "id": "SHOWTASKPANE",
                "type": "ExecuteFunction",
                "name": "Show task pane for add-in"
            },
            {
                "id": "HIDETASKPANE",
                "type": "ExecuteFunction",
                "name": "Hide task pane for add-in"
            }
        ],
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Up"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Down"
                }
            }
        ]
    }
    ```

    有关 JSON 对象详细信息，请参阅 [构造操作](#construct-the-action-objects) 对象和 [构造快捷方式对象](#construct-the-shortcut-objects)。 快捷方式 JSON 的完整架构位于 上的[extended-manifest.schema.js。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

    > [!NOTE]
    > 可以在整个文章中使用"CONTROL"来表示"Ctrl"。

    在稍后的步骤中，操作本身将映射到您编写的函数。 此示例稍后将 SHOWTASKPANE 映射到调用 方法的函数， `Office.addin.showAsTaskpane` 将 HIDETASKPANE 映射到调用 该方法 `Office.addin.hide` 的函数。

## <a name="create-a-mapping-of-actions-to-their-functions"></a>创建操作到其函数的映射

1. 在项目中，打开 元素中的 HTML 页面加载的 JavaScript `<FunctionFile>` 文件。
1. 在 JavaScript 文件中，使用[Office.actions.associate](/javascript/api/office/office.actions#associate) API 将 JSON 文件中指定的每个操作映射到 JavaScript 函数。 将以下 JavaScript 添加到文件中。 关于代码，请注意以下几点。

    - 第一个参数是 JSON 文件的操作之一。
    - 第二个参数是当用户按下映射到 JSON 文件中操作的组合键时运行的函数。

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. 若要继续此示例，请使用 `'SHOWTASKPANE'` 作为第一个参数。
1. 对于函数的正文，使用[Office.addin.showTaskpane](/javascript/api/office/office.addin#showAsTaskpane__)方法打开加载项的任务窗格。 完成后，代码应如下所示：

    ```javascript
    Office.actions.associate('SHOWTASKPANE', function () {
        return Office.addin.showAsTaskpane()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

1. 添加函数的第二个调用，以将操作映射到调用 `Office.actions.associate` `HIDETASKPANE` [Office.addin.hide 的函数](/javascript/api/office/office.addin#hide__)。 示例如下。

    ```javascript
    Office.actions.associate('HIDETASKPANE', function () {
        return Office.addin.hide()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

按照前面的步骤，加载项可通过按 **Ctrl+Alt+Up** 和 **Ctrl+Alt+Down 切换任务窗格的可见性**。 相同的行为显示在 Excel 外接程序[](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)PnP Office中的键盘快捷方式示例GitHub。

## <a name="details-and-restrictions"></a>详细信息和限制

### <a name="construct-the-action-objects"></a>构造操作对象

在上指定对象数组中的对象时 `actions` ，shortcuts.js准则。

- 属性名 `id` 和 `name` 是必需的。
- `id`属性用于唯一标识使用键盘快捷方式调用的操作。
- `name`属性必须是描述操作的用户友好字符串。 它必须是字符 A - Z、a - z、0 - 9 和标点符号"-"、"_"和"+"的组合。
- 属性是可选的。 当前仅 `ExecuteFunction` 支持类型。

示例如下。

```json
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        },
        {
            "id": "HIDETASKPANE",
            "type": "ExecuteFunction",
            "name": "Hide task pane for add-in"
        }
    ]
```

快捷方式 JSON 的完整架构位于 上的[extended-manifest.schema.js。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

### <a name="construct-the-shortcut-objects"></a>构造快捷方式对象

在上指定对象数组中的对象时 `shortcuts` ，shortcuts.js准则。

- 属性名称 `action` 、 `key` 和 `default` 是必需的。
- 该属性的值 `action` 是一个字符串，并且必须与 action 对象 `id` 中的某个属性匹配。
- 该属性 `default` 可以是字符 A - Z、-z、0 - 9 和标点符号"-"、"_"和"+"的任意组合。  (根据惯例，这些属性中不使用小写字母。) 
- 该属性 `default` 必须至少包含 Alt、Ctrl、Shift 和一个 (键的名称) 一个修饰符键。
- Shift 不能用作唯一的修改键。 将 Shift 与 Alt 或 Ctrl 组合使用。
- 对于 Mac，我们还支持 Command 修饰符键。
- 对于 Mac，Alt 映射到 Option 键。 例如Windows命令映射到 Ctrl 键。
- 当两个字符链接到标准键盘中的同一个物理键时，它们是 属性中的同义词;例如，Alt+a 和 Alt+A 是同一快捷方式 `default` ，Ctrl+- 和 Ctrl+ 也是，因为 \_ "-"和"_"是同一个物理键。
- "+"字符指示同时按下其任一侧的键。

示例如下。

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down"
            }
        }
    ]
```

快捷方式 JSON 的完整架构位于 上的[extended-manifest.schema.js。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

> [!NOTE]
> 键提示（也称为连续键快捷方式，例如选择填充颜色的 Excel 快捷方式 **Alt+H、H）** 在加载项中不受Office支持。

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a>避免其他加载项使用组合键

有许多键盘快捷方式已由 Office。 避免为已在使用的外接程序注册键盘快捷方式，但在某些情况下，可能需要替代现有键盘快捷方式或处理已注册同一键盘快捷方式的多个加载项之间的冲突。

如果发生冲突，用户将在第一次尝试使用冲突的键盘快捷方式时看到一个对话框，请注意，此对话框中显示的动作名称是文件中 action 对象中的 属性。 `name` `shortcuts.json`

![插图显示具有单个快捷方式的两个不同操作的冲突模式。](../images/add-in-shortcut-conflict-modal.png)

用户可以选择键盘快捷方式将执行的操作。 做出选择后，保存首选项，供将来使用同一快捷方式。 快捷方式首选项按用户、平台保存。 如果用户希望更改其首选项，他们可以从"告诉我"搜索框中调用"重置Office外接程序快捷方式首选项"命令。  调用命令可清除用户的所有加载项快捷方式首选项，并且用户下次尝试使用冲突快捷方式时，会再次看到冲突对话框提示。

!["告诉我"搜索框显示在Excel快捷方式首选项Office重置操作。](../images/add-in-reset-shortcuts-action.png)

为了获得最佳用户体验，我们建议您尽量减少与这些Excel冲突。

- 请仅使用以下模式的键盘快捷方式： **Ctrl+Shift+Alt+* x***，其中 *x* 是一些其他键。
- 如果您需要更多键盘快捷方式，请检查Excel[键盘](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)快捷方式的列表，并避免在外接程序中使用它们。
- 当键盘焦点位于加载项 UI 内时 **，Ctrl+空格** 键和 **Ctrl+Shift+F10** 将不起作用，因为这些都是基本的辅助功能快捷方式。
- 在 Windows 或 Mac 计算机上，如果"重置 Office 外接程序快捷方式首选项"命令在搜索菜单上不可用，则用户可以通过通过上下文菜单自定义功能区，将该命令手动添加到功能区。

## <a name="customize-the-keyboard-shortcuts-per-platform"></a>自定义每个平台的键盘快捷方式

可以自定义特定于平台的快捷方式。 下面是自定义以下每个平台的快捷方式的对象示例 `shortcuts` `windows` `mac` ：、、。 `web` 请注意，您仍必须具有 `default` 每个快捷方式的快捷键。

在下面的示例中， `default` 键是未指定的任何平台的回退键。 唯一未指定的平台Windows，因此 `default` 该密钥仅适用于Windows。

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up",
                "mac": "Command+Shift+Up",
                "web": "Ctrl+Alt+1",
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down",
                "mac": "Command+Shift+Down",
                "web": "Ctrl+Alt+2"
            }
        }
    ]
```

## <a name="localize-the-keyboard-shortcuts-json"></a>本地化键盘快捷方式 JSON

如果加载项支持多个区域设置，则需要本地化 `name` action 对象的 属性。 此外，如果加载项支持的任何区域设置具有字母或不同的书写系统，因此使用不同的键盘，则你可能还需要本地化快捷方式。 若要了解如何本地化键盘快捷方式 JSON，请参阅 [本地化扩展替代](../develop/localization.md#localize-extended-overrides)。

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>无法重写的浏览器快捷方式

在 Web 上使用自定义键盘快捷方式时，外接程序无法覆盖浏览器所使用的某些键盘快捷方式。此列表是一项正在进行中的工作。 如果发现无法覆盖的其他组合，请使用此页面底部的反馈工具告诉我们。

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="next-steps"></a>后续步骤

- 请参阅[Excel键盘快捷方式](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)示例外接程序。
- 获取有关使用清单的扩展替代 [中的扩展覆盖的概述](../develop/extended-overrides.md)。

---
title: Office 加载项中的自定义键盘快捷方式
description: 了解如何将自定义键盘快捷方式（也称为组合键）添加到 Office 外接程序。
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: dc99674b92ebb415b1d49fb28821d8c2e34c8077
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789147"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>向 Office 外接程序添加自定义键盘快捷方式 (预览) 

键盘快捷方式（也称为组合键）使加载项的用户能够更高效地工作，并且它们通过提供鼠标替代项为残障用户提供外接程序的辅助功能。

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> 若要从已启用键盘快捷方式的工作版本的加载项开始，请克隆并运行 [示例 Excel 键盘快捷方式](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。 准备好向自己的外接程序添加键盘快捷方式后，请继续阅读本文。

向加载项添加键盘快捷方式有三个步骤：

1. [配置加载项的清单](#configure-the-manifest)。
1. [创建或编辑快捷方式 JSON 文件以](#create-or-edit-the-shortcuts-json-file) 定义操作及其键盘快捷方式。
1. [添加](#create-a-mapping-of-actions-to-their-functions) [Office.actions.associate](/javascript/api/office/office.actions#associate) API 的一个或多个运行时调用，以将函数映射到每个操作。

## <a name="configure-the-manifest"></a>配置清单

清单有两个小更改需要更改。 一种是允许加载项使用共享运行时，另一种是指向定义键盘快捷方式的 JSON 格式文件。

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>配置加载项以使用共享运行时

添加自定义键盘快捷方式需要加载项使用共享运行时。 有关详细信息， [请配置外接程序以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

### <a name="link-the-mapping-file-to-the-manifest"></a>将映射文件链接到清单

紧 *(* 清单) 元素的内部，添加 `<VersionOverrides>` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素。 将该属性设置为将在稍后步骤创建的项目中 `Url` JSON 文件的完整 URL。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>创建或编辑快捷方式 JSON 文件

在项目中创建 JSON 文件。 确保文件的路径与为 `Url` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素的属性指定的位置匹配。 此文件将描述键盘快捷方式，以及这些快捷方式将调用的操作。

1. 在 JSON 文件中，有两个数组。 操作数组将包含定义要调用的操作的对象，快捷方式数组将包含将组合键映射到操作的对象。 如以下示例所示：

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
                    "default": "CTRL+SHIFT+UP"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+DOWN"
                }
            }
        ]
    }
    ```

    有关 JSON 对象详细信息，请参阅 [构造操作](#constructing-the-action-objects) 对象和 [构造快捷方式对象](#constructing-the-shortcut-objects)。 快捷方式 JSON 的完整架构位于extended-manifest.schema.js[on。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

    > [!NOTE]
    > 可以在整个文章中使用"CONTROL"来表示"CTRL"。

    在稍后的步骤中，操作本身将映射到您编写的函数。 此示例稍后将 SHOWTASKPANE 映射到调用该方法的函数， `Office.addin.showAsTaskpane` 将 HIDETASKPANE 映射到调用该方法 `Office.addin.hide` 的函数。

## <a name="create-a-mapping-of-actions-to-their-functions"></a>创建操作到其函数的映射

1. 在项目中，打开由元素中的 HTML 页面加载的 JavaScript `<FunctionFile>` 文件。
1. 在 JavaScript 文件中，使用 [Office.actions.associate](/javascript/api/office/office.actions#associate) API 将 JSON 文件中指定的每个操作映射到 JavaScript 函数。 将以下 JavaScript 添加到文件。 关于代码，请注意以下几点：

    - 第一个参数是 JSON 文件的操作之一。
    - 第二个参数是当用户按下映射到 JSON 文件中操作的组合键时运行的函数。

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. 若要继续该示例，请使用 `'SHOWTASKPANE'` 作为第一个参数。
1. 对于函数的正文，使用 [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) 方法打开加载项的任务窗格。 完成后，代码应如下所示：

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

1. 添加函数的第二个调用，以将操作映射到调用 `Office.actions.associate` `HIDETASKPANE` [Office.addin.hide 的函数](/javascript/api/office/office.addin#hide--)。 示例如下：

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

按照前面的步骤，加载项可通过按 **Ctrl+Shift+向上** 键和 **Ctrl+Shift+向下** 箭头键切换任务窗格的可见性。 这是与示例 excel 键盘快捷方式外接程序 [中显示的相同行为](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。

## <a name="details-and-restrictions"></a>详细信息和限制

### <a name="constructing-the-action-objects"></a>构造操作对象

指定对象数组中的对象时，请使用以下 `action` shortcuts.js：

- 属性名称 `id` 且 `name` 是必需的。
- 该属性 `id` 用于唯一标识使用键盘快捷方式调用的操作。
- 该属性 `name` 必须是描述操作的用户友好字符串。 它必须是字符 A - Z、a - z、0 - 9 和标点符号"-"、"_"和"+"的组合。
- 属性是可选的。 当前 `ExecuteFunction` 仅支持类型。

示例如下：

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

快捷方式 JSON 的完整架构位于extended-manifest.schema.js[on。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

### <a name="constructing-the-shortcut-objects"></a>构造快捷方式对象

指定对象数组中的对象时，请使用以下 `shortcuts` shortcuts.js：

- 属性名称 `action` ， `key` 和 `default` 是必需的。
- 该属性的值 `action` 是一个字符串，并且必须与 action 对象 `id` 中的一个属性匹配。
- 该属性 `default` 可以是字符 A - Z、-z、0 - 9 和标点符号"-"、"_"和"+"的任意组合。  (根据惯例，这些属性中不使用小写字母。) 
- 该属性必须包含至少一个修饰符键的名称 (`default` Alt、Ctrl、Shift) 一个其他键。
- 对于 Mac，我们还支持 COMMAND 修饰符键。
- 对于 Mac，ALT 映射到 OPTION 键。 对于 Windows，COMMAND 映射到 Ctrl 键。
- 当两个字符链接到标准键盘中的同一物理键时，它们是属性中的同义词;例如，Alt+a 和 Alt+A 是同一快捷方式 `default` ，Ctrl+- 和 Ctrl+ 也是，因为"-"和"_"是同一物理键。 \_
- "+"字符指示同时按下其任一侧的键。

示例如下：

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+UP"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "CTRL+SHIFT+DOWN"
            }
        }
    ]
```

快捷方式 JSON 的完整架构位于extended-manifest.schema.js[on。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

> [!NOTE]
> Office 加载项不支持键提示（也称为顺序键快捷方式，如 Excel 快捷方式选择填充颜色 **Alt+H、H）。**

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>当焦点位于任务窗格中时，使用快捷方式

目前，只有在用户焦点位于工作表中时，才能调用 Office 加载项的键盘快捷方式。 当用户的焦点位于 Office UI (（如任务窗格) ）中时，不会忽略任何加载项的快捷方式。 作为一种解决方法，加载项可以定义键盘处理程序，当用户的焦点位于加载项 UI 内时，可以调用某些操作。

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>使用 Office 或其他外接程序已使用的键组合

在预览期间，系统无法确定当用户按下由加载项以及 Office 或其他加载项注册的键组合时会发生什么情况。 行为未定义。

目前，当两个或多个加载项已注册相同的键盘快捷方式时，没有解决方法，但您可以通过这些好的做法最大程度地减少与 Excel 的冲突：

- 在外接程序中，仅使用以下模式的键盘快捷方式：**Ctrl+Shift+Alt+* x***，其中 *x* 是一些其他键。
- 如果您需要更多键盘快捷方式，请检查 [Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)键盘快捷方式的列表，并避免在加载项中使用它们。

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>无法重写的浏览器快捷方式

不能使用下列任何键盘组合。 它们由浏览器使用，不能重写。 此列表是一项正在进行中的工作。 如果发现无法替代的其他组合，请使用此页面底部的反馈工具告知我们。

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="next-steps"></a>后续步骤

- 请参阅示例加载项[excel-keyboard-shortcuts。](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)

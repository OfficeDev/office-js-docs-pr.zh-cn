---
title: Office 外接程序中的自定义键盘快捷方式
description: 了解如何将自定义键盘快捷方式（也称为键组合）添加到 Office 外接程序。
ms.date: 11/09/2020
localization_priority: Normal
ms.openlocfilehash: 40009dd92787b7c220bb8cfc741cffb2e4b68a9e
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132037"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>将自定义键盘快捷方式添加到 Office 外接 (预览) 

键盘快捷方式（也称为键组合）使您的外接程序的用户可以更高效地工作，并通过提供鼠标替换功能为残障用户改进了加载项的辅助功能。

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> 若要从已启用的键盘快捷方式开始使用加载项的工作版本，请克隆并运行示例 [Excel 键盘快捷方式](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。 准备好将键盘快捷方式添加到自己的外接程序后，请继续阅读本文。

将键盘快捷方式添加到外接程序中有三个步骤：

1. [配置加载项的清单](#configure-the-manifest)。
1. [创建或编辑快捷方式 JSON 文件](#create-or-edit-the-shortcuts-json-file) 以定义操作及其键盘快捷方式。
1. [添加一个或多个 Office 的运行时调用](#create-a-mapping-of-actions-to-their-functions) [。关联](/javascript/api/office/office.actions#associate) API 以将某个函数映射到每个操作。

## <a name="configure-the-manifest"></a>配置清单

对清单进行了两处较小的更改。 一种是使外接程序能够使用共享运行时，而另一种是指向您定义了键盘快捷方式的 JSON 格式的文件。

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>将加载项配置为使用共享运行时

若要添加自定义键盘快捷方式，您的加载项需要使用共享运行时。 有关详细信息，请 [配置外接程序以使用共享运行时](../excel/configure-your-add-in-to-use-a-shared-runtime.md)。

### <a name="link-the-mapping-file-to-the-manifest"></a>将映射文件链接到清单

在 *下面* 紧接着 (不在 `<VersionOverrides>` 清单中的元素) 元素中，添加一个 [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素。 将 `Url` 属性设置为项目中您将在后续步骤中创建的 JSON 文件的完整 URL。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>创建或编辑快捷方式 JSON 文件

在项目中创建一个 JSON 文件。 确保文件的路径与您为 ExtendedOverrides 元素的属性指定的位置相匹配 `Url` 。 [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 此文件将介绍你的键盘快捷方式以及它们将调用的操作。

1. 在 JSON 文件中，有两个数组。 操作数组将包含定义要调用的操作的对象，并且快捷键数组将包含将键组合映射到操作的对象。 如以下示例所示：

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

    有关 JSON 对象的详细信息，请参阅 [构造 action 对象](#constructing-the-action-objects) 和 [构造快捷方式对象](#constructing-the-shortcut-objects)。 快捷键 JSON 的完整架构位于 [extended-manifest.schema.js](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。  (注意：指向架构的链接可能在预览周期中的早期阶段无法运行。 ) 

    > [!NOTE]
    > 在本文中，您可以使用 "控制" 代替 "CTRL"。

    在后续步骤中，操作本身将映射到您编写的函数。 在此示例中，您稍后会将 SHOWTASKPANE 映射到一个函数，该函数调用 `Office.addin.showAsTaskpane` 方法和 HIDETASKPANE 到调用该 `Office.addin.hide` 方法的函数。

## <a name="create-a-mapping-of-actions-to-their-functions"></a>创建操作到它们的函数的映射

1. 在您的项目中，打开元素中的 HTML 页面加载的 JavaScript 文件 `<FunctionFile>` 。
1. 在 JavaScript 文件中，使用 [Office. 操作。关联](/javascript/api/office/office.actions#associate) API 将您在 JSON 文件中指定的每个操作映射到一个 JavaScript 函数。 向文件中添加以下 JavaScript。 请注意有关代码的以下内容：

    - 第一个参数是 JSON 文件中的一项操作。
    - 第二个参数是当用户按下将映射到 JSON 文件中的操作的组合键时运行的函数。

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. 若要继续本示例，请使用 `'SHOWTASKPANE'` 作为第一个参数。
1. 对于函数的主体，请使用 [showTaskpane](/javascript/api/office/office.addin#showastaskpane--) 方法打开外接程序的任务窗格。 完成后，代码应类似于以下内容：

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

1. 添加第二个函数调用， `Office.actions.associate` 以将 `HIDETASKPANE` 操作映射到一个调用了 [.addin](/javascript/api/office/office.addin#hide--)的函数。 示例如下：

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

按照前面的步骤，你的外接程序可以通过按 **ctrl + shift + 向上箭头键** 和 **Ctrl + Shift + 向下箭头键** 来切换任务窗格的可见性。 这与 [示例 excel 键盘快捷方式加载项](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)中所示的行为相同。

## <a name="details-and-restrictions"></a>详细信息和限制

### <a name="constructing-the-action-objects"></a>构造 action 对象

在中指定 shortcuts.js数组中的对象时，请使用以下准则 `action` ：

- 属性名称 `id` ，并且 `name` 是强制性的。
- 该 `id` 属性用于唯一标识要使用键盘快捷方式调用的操作。
- 该 `name` 属性必须是描述操作的用户友好字符串。 它必须是字符 a-z、a-z、0-9 和标点符号 "-"、"_" 和 "+" 的组合。
- 属性是可选的。 目前仅 `ExecuteFunction` 支持类型。

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

快捷键 JSON 的完整架构位于 [extended-manifest.schema.js](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。  (注意：指向架构的链接可能在预览周期中的早期阶段无法运行。 ) 

### <a name="constructing-the-shortcut-objects"></a>构造快捷方式对象

在中指定 shortcuts.js数组中的对象时，请使用以下准则 `shortcuts` ：

- 属性名称 `action` 、 `key` 和 `default` 是必需的。
- 该属性的值 `action` 是一个字符串，并且必须与 `id` action 对象中的一个属性相匹配。
- 该 `default` 属性可以是字符 a-z、a-z、0-9 和标点符号 "-"、"_" 和 "+" 的任意组合。  (按惯例，在这些属性中不使用小写字母。 ) 
- `default`属性必须包含至少一个修改键的名称 (ALT、CTRL、SHIFT) 且仅包含一个其他键。
- 对于 Mac，我们还支持命令修改键。
- 对于 Mac，将 ALT 映射到选项键。 对于 Windows，命令映射到 CTRL 键。
- 当两个字符链接到标准键盘中的同一个物理键时，它们就是属性中的同义词 `default` ; 例如，alt + a 和 alt + a 是相同的快捷方式，因此是 ctrl +-和 ctrl +， \_ 因为 "-" 和 "_" 是相同的物理键。
- "+" 字符指示同时按下的键的任意一侧的键。

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

快捷键 JSON 的完整架构位于 [extended-manifest.schema.js](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。  (注意：指向架构的链接可能在预览周期中的早期阶段无法运行。 ) 

> [!NOTE]
> 快捷键提示（也称为连续键快捷方式，例如，用于选择填充颜色的 Excel 快捷方式 **Alt + h，h**）在 Office 加载项中不受支持。

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>当焦点在任务窗格中时使用快捷方式

目前，只有当用户的焦点在工作表中时，才能调用 Office 外接程序的键盘快捷方式。 当用户的焦点位于 Office UI (（例如任务窗格) ）中时，不会忽略任何加载项的快捷方式。 作为一种解决方法，加载项可以定义键盘处理程序，当用户的焦点位于外接程序 UI 中时，可以调用某些操作。

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>使用已由 Office 或其他加载项使用的组合键

在预览期间，没有系统可用于确定当用户按外接程序注册的组合键以及由 Office 或其他外接程序注册时，会发生什么情况。 行为未定义。

目前，如果两个或更多个加载项注册了相同的键盘快捷方式，但您可以最大限度地减少与 Excel 的冲突，请使用以下这些好的做法：

- 在外接程序中仅使用具有以下模式的键盘快捷方式： **Ctrl + Shift + Alt +* x * * *，其中 *x* 是另一个键。
- 如果需要更多键盘快捷方式，请查看 [Excel 键盘快捷方式列表](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)，并避免在外接程序中使用其中任何一个。

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>无法覆盖的浏览器快捷方式

您不能使用以下任何键盘组合。 它们由浏览器使用，不能覆盖。 此列表是一项正在进行的工作。 如果发现无法覆盖的其他组合，请使用本页底部的反馈工具告知我们。

- Ctrl + N
- Ctrl + Shift + N
- Ctrl + T
- Ctrl + Shift + T
- Ctrl + W
- Ctrl + PgUp/PgDn

## <a name="next-steps"></a>后续步骤

- 请参阅示例加载项 [excel-键盘快捷方式](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。

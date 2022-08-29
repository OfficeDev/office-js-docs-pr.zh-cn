---
title: Office 加载项中的自定义键盘快捷方式
description: 了解如何将自定义键盘快捷方式（也称为键组合）添加到 Office 外接程序。
ms.date: 11/22/2021
localization_priority: Normal
ms.openlocfilehash: 462e5bfdd4e7f825318d6affb631beafc7c08fe5
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423018"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a>将自定义键盘快捷方式添加到 Office 加载项

键盘快捷方式（也称为键组合）使外接程序的用户能够更高效地工作。 键盘快捷方式还通过提供鼠标的替代方法，提高残障用户的外接程序的辅助功能。

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> 若要从已启用键盘快捷方式的外接程序的工作版本开始，请克隆并运行示例 [Excel 键盘快捷方式](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts)。 准备好向自己的外接程序添加键盘快捷方式时，请继续阅读本文。

向加载项添加键盘快捷方式有三个步骤。

1. [配置加载项的清单](#configure-the-manifest)。
1. [创建或编辑快捷方式 JSON 文件](#create-or-edit-the-shortcuts-json-file) 以定义操作及其键盘快捷方式。
1. 添加 [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) API 的[一个或多个运行时调用](#create-a-mapping-of-actions-to-their-functions)，以将函数映射到每个操作。

## <a name="configure-the-manifest"></a>配置清单

要对清单进行两个小更改。 一是使外接程序能够使用共享运行时，另一种是指向定义键盘快捷方式的 JSON 格式的文件。

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>将外接程序配置为使用共享运行时

添加自定义键盘快捷方式需要外接程序使用 [共享运行时](../testing/runtimes.md#shared-runtime)。 有关详细信息，请 [将加载项配置为使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

### <a name="link-the-mapping-file-to-the-manifest"></a>将映射文件链接到清单

紧 *接着* (不在清单中元素) **\<VersionOverrides\>** 内，添加 [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) 元素。 将属性 `Url` 设置为项目中将在后续步骤中创建的 JSON 文件的完整 URL。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>创建或编辑快捷方式 JSON 文件

在项目中创建 JSON 文件。 请确保文件的路径与为 `Url` [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) 元素的属性指定的位置匹配。 此文件将描述键盘快捷方式及其将调用的操作。

1. 在 JSON 文件中，有两个数组。 操作数组将包含定义要调用的操作的对象，快捷方式数组将包含将键组合映射到操作上的对象。 以下是示例。
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

    有关 JSON 对象的详细信息，请参阅 [构造操作对象](#construct-the-action-objects) 和 [构造快捷方式对象](#construct-the-shortcut-objects)。 快捷方式 JSON 的完整架构位于 [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json) 中。

    > [!NOTE]
    > 可在本文中使用“CONTROL”代替“Ctrl”。

    在后面的步骤中，操作本身将映射到你编写的函数。 在此示例中，稍后会将 SHOWTASKPANE 映射到调用 `Office.addin.showAsTaskpane` 方法的函数，并将 HIDETASKPANE 映射到调用该方法的 `Office.addin.hide` 函数。

## <a name="create-a-mapping-of-actions-to-their-functions"></a>创建操作到其函数的映射

1. 在项目中，打开元素中由 HTML 页面加载的 **\<FunctionFile\>** JavaScript 文件。
1. 在 JavaScript 文件中，使用 [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) API 将 JSON 文件中指定的每个操作映射到 JavaScript 函数。 将以下 JavaScript 添加到文件中。 请注意以下代码。

    - 第一个参数是 JSON 文件中的操作之一。
    - 第二个参数是当用户按映射到 JSON 文件中的操作的键组合时运行的函数。

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. 若要继续该示例，请用作 `'SHOWTASKPANE'` 第一个参数。
1. 对于函数的正文，请使用 [Office.addin.showAsTaskpane](/javascript/api/office/office.addin#office-office-addin-showastaskpane-member(1)) 方法打开加载项的任务窗格。 完成后，代码应如下所示：

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

1. 添加第二次函数 `Office.actions.associate` 调用以将操作映射 `HIDETASKPANE` 到调用 [Office.addin.hide](/javascript/api/office/office.addin#office-office-addin-hide-member(1)) 的函数。 示例如下。

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

按照前面的步骤，外接程序可以通过按 **Ctrl+Alt+Up** 和 **Ctrl+Alt+Down** 来切换任务窗格的可见性。 在 GitHub 中的 Office 外接程序 PnP 存储库的 [Excel 键盘快捷方式](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) 示例中显示了相同的行为。

## <a name="details-and-restrictions"></a>详细信息和限制

### <a name="construct-the-action-objects"></a>构造操作对象

在 shortcuts.json 数组中指定对象时， `actions` 请使用以下准则。

- 属性名称 `id` 是 `name` 必需的。
- 该 `id` 属性用于唯一标识使用键盘快捷方式调用的操作。
- 该 `name` 属性必须是描述操作的用户友好字符串。 它必须是字符 A - Z、a - z、0 - 9 和标点符号“-”、“_”和“+”的组合。
- 属性是可选的。 目前仅 `ExecuteFunction` 支持类型。

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

快捷方式 JSON 的完整架构位于 [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json) 中。

### <a name="construct-the-shortcut-objects"></a>构造快捷方式对象

在 shortcuts.json 数组中指定对象时， `shortcuts` 请使用以下准则。

- 属性名称 `action`， `key`并且 `default` 是必需的。
- 该属性的 `action` 值是一个字符串，必须与操作对象中的某个 `id` 属性匹配。
- 该 `default` 属性可以是字符 A - Z、a-z、0 - 9 和标点符号“-”、“_”和“+”的任何组合。  (根据约定，这些属性中不使用小写字母。) 
- 该 `default` 属性必须包含至少一个修饰符键的名称 (Alt、Ctrl、Shift) ，并且仅包含另一个键。
- 不能将 Shift 用作唯一的修饰符键。 将 Shift 与 Alt 或 Ctrl 合并。
- 对于 Mac，我们还支持命令修饰符密钥。
- 对于 Mac，Alt 将映射到“选项”键。 对于 Windows，命令映射到 Ctrl 键。
- 当两个字符链接到标准键盘中的同一物理键时，它们是属性中的 `default` 同义词;例如，Alt+a 和 Alt+A 是相同的快捷方式，Ctrl+和 Ctrl+\_ 也是如此，因为“-”和“_”是相同的物理键。
- “+”字符指示同时按下其两侧的键。

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

快捷方式 JSON 的完整架构位于 [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json) 中。

> [!NOTE]
> Office 外接程序不支持键提示（也称为顺序键快捷方式），例如用于选择填充颜色 **Alt+H、H** 的 Excel 快捷方式。

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a>避免其他加载项使用的密钥组合

Office 已使用许多键盘快捷方式。 避免为已在使用的外接程序注册键盘快捷方式，但在某些情况下，可能需要重写现有键盘快捷方式或处理已注册同一键盘快捷方式的多个加载项之间的冲突。

如果出现冲突，用户会在首次尝试使用冲突键盘快捷方式时看到对话框。 请注意，此对话框中显示的加载项选项的文本来自 `name` 文件中的操作对象中的 `shortcuts.json` 属性。

![显示一个冲突模式的插图，其中包含单个快捷方式的两个不同的操作。](../images/add-in-shortcut-conflict-modal.png)

用户可以选择键盘快捷方式将执行的操作。 进行选择后，将保存首选项以供将来使用相同的快捷方式。 每个用户（每个平台）保存快捷方式首选项。 如果用户希望更改首选项，则可以从“**告诉我**”搜索框调用 **“重置 Office 加载项快捷方式首选项**”命令。 调用命令可清除用户的所有加载项快捷方式首选项，下次尝试使用冲突快捷方式时，系统会再次提示用户使用冲突对话框。

![Excel 中的“告诉我”搜索框，显示“重置 Office 外接程序快捷方式首选项”操作。](../images/add-in-reset-shortcuts-action.png)

为了获得最佳用户体验，我们建议将与 Excel 的冲突与这些良好做法最小化。

- 使用以下模式的键盘快捷方式：**Ctrl+Shift+Alt+* x***，其中 *x* 是一些其他键。
- 如果需要更多键盘快捷方式，请检查 [Excel 键盘快捷方式列表](https://support.microsoft.com/office/1798d9d5-842a-42b8-9c99-9b7213f0040f)，避免在外接程序中使用任何键盘快捷方式。
- 当键盘焦点位于加载项 UI 内时， **Ctrl+空格键** 和 **Ctrl+Shift+F10** 将不起作用，因为这些是必不可少的辅助功能快捷方式。
- 在 Windows 或 Mac 计算机上，如果搜索菜单上没有“重置 Office 加载项快捷方式首选项”命令，用户可以通过上下文菜单自定义功能区，手动将命令添加到功能区。

## <a name="customize-the-keyboard-shortcuts-per-platform"></a>自定义每个平台的键盘快捷方式

可以自定义特定于平台的快捷方式。 下面是自定义以下每个平台的快捷方式的对象示例`shortcuts`： `windows`， ， `mac``web`。 请注意，每个快捷方式仍必须有 `default` 快捷键。

在以下示例中 `default` ，密钥是未指定的任何平台的回退键。 唯一未指定的平台是 Windows，因此该 `default` 密钥将仅应用于 Windows。

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

如果外接程序支持多个区域设置，则需要本地化 `name` 操作对象的属性。 此外，如果外接程序支持的任何区域设置具有不同的字母表或写入系统，因此键盘也不同，则可能需要本地化快捷方式。 有关如何本地化键盘快捷方式 JSON 的信息，请参阅 [Localize 扩展重写](../develop/localization.md#localize-extended-overrides)。

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>无法重写的浏览器快捷方式

在 Web 上使用自定义键盘快捷方式时，加载项无法重写浏览器使用的某些键盘快捷方式。此列表是正在进行的工作。 如果发现无法重写的其他组合，请使用此页面底部的反馈工具告知我们。

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="enable-custom-keyboard-shortcuts-for-specific-users"></a>为特定用户启用自定义键盘快捷方式

你的外接程序可以让用户将加载项的操作重新分配到备用键盘组合。

> [!NOTE]
> 本部分中所述的 API 需要 [设置 KeyboardShortcuts 1.1](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets) 要求。

使用 [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) 方法将用户的自定义键盘组合分配给加载项操作。 该方法采用类型 `{[actionId:string]: string|null}`参数，其中 `actionId`s 是必须在外接程序的扩展清单 JSON 中定义的操作 ID 的子集。 这些值是用户首选的键组合。 该值也可以是 `null`，它将删除任何自定义项， `actionId` 并还原到外接程序的扩展清单 JSON 中定义的默认键盘组合。

如果用户登录到 Office，则自定义组合将保存在每个平台用户的漫游设置中。 匿名用户目前不支持自定义快捷方式。

```javascript
const userCustomShortcuts = {
    SHOWTASKPANE:"CTRL+SHIFT+1", 
    HIDETASKPANE:"CTRL+SHIFT+2"
};
Office.actions.replaceShortcuts(userCustomShortcuts)
    .then(function () {
        console.log("Successfully registered.");
    })
    .catch(function (ex) {
        if (ex.code == "InvalidOperation") {
            console.log("ActionId does not exist or shortcut combination is invalid.");
        }
    });
```

若要了解用户已使用的快捷方式，请调用 [Office.actions.getShortcuts](/javascript/api/office/office.actions#office-office-actions-getshortcuts-member) 方法。 此方法返回一个类型的 `[actionId:string]:string|null}`对象，其中值表示用户调用指定操作时必须使用的当前键盘组合。 这些值可以来自三个不同的源：

- 如果与快捷方式发生冲突，并且用户已选择使用其他操作 (本机或其他加载项) 用于该键盘组合，则返回的值将是 `null` 因为快捷方式已被重写，并且用户当前无法使用键盘组合来调用该外接程序操作。
- 如果已使用 [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) 方法自定义快捷方式，则返回的值将是自定义键盘组合。
- 如果快捷方式尚未重写或自定义，它将从外接程序的扩展清单 JSON 返回值。

示例如下。

```javascript
Office.actions.getShortcuts()
    .then(function (userShortcuts) {
       for (const action in userShortcuts) {
           let shortcut = userShortcuts[action];
           console.log(action + ": " + shortcut);
       }
    });

```

如《 [避免其他加载项使用的关键组合](#avoid-key-combinations-in-use-by-other-add-ins)》中所述，最好避免快捷方式中的冲突。 若要发现一个或多个键组合是否已在使用中，请将它们作为字符串数组传递给 [Office.actions.areShortcutsInUse](/javascript/api/office/office.actions#office-office-actions-areshortcutsinuse-member) 方法。 该方法返回一个报表，其中包含已以类型 `{shortcut: string, inUse: boolean}`对象数组的形式使用的键组合。 该 `shortcut` 属性是关键组合，例如“CTRL+SHIFT+1”。 如果组合已注册到另一个操作，则 `inUse` 属性设置为 `true`。 例如，`[{shortcut: "CTRL+SHIFT+1", inUse: true}, {shortcut: "CTRL+SHIFT+2", inUse: false}]`。 下面的代码片段就是一个示例：

```javascript
const shortcuts = ["CTRL+SHIFT+1", "CTRL+SHIFT+2"];
Office.actions.areShortcutsInUse(shortcuts)
    .then(function (inUseArray) {
        const availableShortcuts = inUseArray.filter(function (shortcut) { return !shortcut.inUse; });
        console.log(availableShortcuts);
        const usedShortcuts = inUseArray.filter(function (shortcut) { return shortcut.inUse; });
        console.log(usedShortcuts);
    });

```

## <a name="next-steps"></a>后续步骤

- 请参阅 [Excel 键盘快捷方式](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) 示例加载项。
- 获取有关在 [使用清单的扩展替代的 Work](../develop/extended-overrides.md) 中使用扩展替代的概述。

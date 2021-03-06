---
title: Office 加载项中的自定义键盘快捷方式
description: 了解如何将自定义键盘快捷方式（也称为组合键）添加到 Office 外接程序。
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: c767c6d5bc23f0a44422452839cd8bdf87bd8715
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505197"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a><span data-ttu-id="55ecb-103">向 Office 外接程序添加自定义键盘快捷方式 (预览) </span><span class="sxs-lookup"><span data-stu-id="55ecb-103">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>

<span data-ttu-id="55ecb-104">键盘快捷方式（也称为组合键）使加载项的用户能够更高效地工作，并且它们通过提供鼠标的替代项为残障用户改进外接程序的辅助功能。</span><span class="sxs-lookup"><span data-stu-id="55ecb-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="55ecb-105">若要从已启用键盘快捷方式的加载项的工作版本开始，请克隆并运行 [示例 Excel 键盘快捷方式](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。</span><span class="sxs-lookup"><span data-stu-id="55ecb-105">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="55ecb-106">准备好向自己的外接程序添加键盘快捷方式后，请继续阅读本文。</span><span class="sxs-lookup"><span data-stu-id="55ecb-106">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="55ecb-107">向加载项添加键盘快捷方式有三个步骤：</span><span class="sxs-lookup"><span data-stu-id="55ecb-107">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="55ecb-108">[配置加载项的清单](#configure-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="55ecb-108">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="55ecb-109">[创建或编辑快捷方式 JSON 文件](#create-or-edit-the-shortcuts-json-file) 以定义操作及其键盘快捷方式。</span><span class="sxs-lookup"><span data-stu-id="55ecb-109">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="55ecb-110">[添加](#create-a-mapping-of-actions-to-their-functions) [Office.actions.associate](/javascript/api/office/office.actions#associate) API 的一个或多个运行时调用，以将函数映射到每个操作。</span><span class="sxs-lookup"><span data-stu-id="55ecb-110">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="55ecb-111">配置清单</span><span class="sxs-lookup"><span data-stu-id="55ecb-111">Configure the manifest</span></span>

<span data-ttu-id="55ecb-112">清单有两个小更改需要进行。</span><span class="sxs-lookup"><span data-stu-id="55ecb-112">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="55ecb-113">一种是允许外接程序使用共享运行时，另一种是指向定义键盘快捷方式的 JSON 格式文件。</span><span class="sxs-lookup"><span data-stu-id="55ecb-113">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="55ecb-114">配置外接程序以使用共享运行时</span><span class="sxs-lookup"><span data-stu-id="55ecb-114">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="55ecb-115">添加自定义键盘快捷方式需要加载项使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="55ecb-115">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="55ecb-116">有关详细信息， [请配置外接程序以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="55ecb-116">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="55ecb-117">将映射文件链接到清单</span><span class="sxs-lookup"><span data-stu-id="55ecb-117">Link the mapping file to the manifest</span></span>

<span data-ttu-id="55ecb-118">紧 *(* 清单) 元素的内部，添加 `<VersionOverrides>` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="55ecb-118">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="55ecb-119">将该属性设置为项目中将在稍后步骤创建的 `Url` JSON 文件的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="55ecb-119">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="55ecb-120">创建或编辑快捷方式 JSON 文件</span><span class="sxs-lookup"><span data-stu-id="55ecb-120">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="55ecb-121">在项目中创建 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="55ecb-121">Create a JSON file in your project.</span></span> <span data-ttu-id="55ecb-122">确保文件的路径与为 `Url` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素的属性指定的位置匹配。</span><span class="sxs-lookup"><span data-stu-id="55ecb-122">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="55ecb-123">此文件将描述键盘快捷方式，以及这些快捷方式将调用的操作。</span><span class="sxs-lookup"><span data-stu-id="55ecb-123">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="55ecb-124">在 JSON 文件中，有两个数组。</span><span class="sxs-lookup"><span data-stu-id="55ecb-124">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="55ecb-125">操作数组将包含定义要调用的操作的对象，快捷方式数组将包含将组合键映射到操作的对象。</span><span class="sxs-lookup"><span data-stu-id="55ecb-125">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="55ecb-126">如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="55ecb-126">Here is an example:</span></span>

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

    <span data-ttu-id="55ecb-127">有关 JSON 对象详细信息，请参阅 [构造操作](#constructing-the-action-objects) 对象 [和构造快捷方式对象](#constructing-the-shortcut-objects)。</span><span class="sxs-lookup"><span data-stu-id="55ecb-127">For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span></span> <span data-ttu-id="55ecb-128">快捷方式 JSON 的完整架构位于extended-manifest.schema.js[on。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="55ecb-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="55ecb-129">您可以使用"CONTROL"来表示本文中的"CTRL"。</span><span class="sxs-lookup"><span data-stu-id="55ecb-129">You can use "CONTROL" in place of "CTRL" throughout this article.</span></span>

    <span data-ttu-id="55ecb-130">在稍后的步骤中，操作本身将映射到您编写的函数。</span><span class="sxs-lookup"><span data-stu-id="55ecb-130">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="55ecb-131">此示例稍后将 SHOWTASKPANE 映射到调用该方法的函数， `Office.addin.showAsTaskpane` 将 HIDETASKPANE 映射到调用该方法 `Office.addin.hide` 的函数。</span><span class="sxs-lookup"><span data-stu-id="55ecb-131">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="55ecb-132">创建操作到其函数的映射</span><span class="sxs-lookup"><span data-stu-id="55ecb-132">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="55ecb-133">在项目中，打开 HTML 页面在元素中加载的 JavaScript `<FunctionFile>` 文件。</span><span class="sxs-lookup"><span data-stu-id="55ecb-133">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="55ecb-134">在 JavaScript 文件中，使用 [Office.actions.associate](/javascript/api/office/office.actions#associate) API 将 JSON 文件中指定的每个操作映射到 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="55ecb-134">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="55ecb-135">将以下 JavaScript 添加到文件中。</span><span class="sxs-lookup"><span data-stu-id="55ecb-135">Add the following JavaScript to the file.</span></span> <span data-ttu-id="55ecb-136">有关代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="55ecb-136">Note the following about the code:</span></span>

    - <span data-ttu-id="55ecb-137">第一个参数是 JSON 文件的操作之一。</span><span class="sxs-lookup"><span data-stu-id="55ecb-137">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="55ecb-138">第二个参数是在用户按下映射到 JSON 文件中操作的组合键时运行的函数。</span><span class="sxs-lookup"><span data-stu-id="55ecb-138">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="55ecb-139">若要继续该示例，请使用 `'SHOWTASKPANE'` 作为第一个参数。</span><span class="sxs-lookup"><span data-stu-id="55ecb-139">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="55ecb-140">对于函数的正文，使用 [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) 方法打开加载项的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="55ecb-140">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="55ecb-141">完成后，代码应如下所示：</span><span class="sxs-lookup"><span data-stu-id="55ecb-141">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="55ecb-142">添加第二个函数调用，以将操作映射到调用 `Office.actions.associate` `HIDETASKPANE` [Office.addin.hide 的函数](/javascript/api/office/office.addin#hide--)。</span><span class="sxs-lookup"><span data-stu-id="55ecb-142">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="55ecb-143">示例如下：</span><span class="sxs-lookup"><span data-stu-id="55ecb-143">The following is an example:</span></span>

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

<span data-ttu-id="55ecb-144">按照上述步骤，加载项可通过按 **Ctrl+Shift+向上** 箭头键和 **Ctrl+Shift+向下** 箭头键来切换任务窗格的可见性。</span><span class="sxs-lookup"><span data-stu-id="55ecb-144">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**.</span></span> <span data-ttu-id="55ecb-145">这是与示例 excel 键盘快捷方式加载项 [中显示的相同行为](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。</span><span class="sxs-lookup"><span data-stu-id="55ecb-145">This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="55ecb-146">详细信息和限制</span><span class="sxs-lookup"><span data-stu-id="55ecb-146">Details and restrictions</span></span>

### <a name="constructing-the-action-objects"></a><span data-ttu-id="55ecb-147">构造操作对象</span><span class="sxs-lookup"><span data-stu-id="55ecb-147">Constructing the action objects</span></span>

<span data-ttu-id="55ecb-148">当指定对象数组中的对象时， `action` 请使用以下shortcuts.js：</span><span class="sxs-lookup"><span data-stu-id="55ecb-148">Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:</span></span>

- <span data-ttu-id="55ecb-149">属性名称 `id` 是 `name` 必需属性。</span><span class="sxs-lookup"><span data-stu-id="55ecb-149">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="55ecb-150">`id`该属性用于唯一标识使用键盘快捷方式调用的操作。</span><span class="sxs-lookup"><span data-stu-id="55ecb-150">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="55ecb-151">该属性 `name` 必须是描述操作的用户友好字符串。</span><span class="sxs-lookup"><span data-stu-id="55ecb-151">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="55ecb-152">它必须是字符 A - Z、a - z、0 - 9 以及标点符号"-"、"_"和"+"的组合。</span><span class="sxs-lookup"><span data-stu-id="55ecb-152">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="55ecb-153">属性是可选的。</span><span class="sxs-lookup"><span data-stu-id="55ecb-153">The `type` property is optional.</span></span> <span data-ttu-id="55ecb-154">当前仅 `ExecuteFunction` 支持类型。</span><span class="sxs-lookup"><span data-stu-id="55ecb-154">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="55ecb-155">示例如下：</span><span class="sxs-lookup"><span data-stu-id="55ecb-155">The following is an example:</span></span>

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

<span data-ttu-id="55ecb-156">快捷方式 JSON 的完整架构位于extended-manifest.schema.js[on。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="55ecb-156">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="constructing-the-shortcut-objects"></a><span data-ttu-id="55ecb-157">构造快捷方式对象</span><span class="sxs-lookup"><span data-stu-id="55ecb-157">Constructing the shortcut objects</span></span>

<span data-ttu-id="55ecb-158">指定对象数组中的对象时， `shortcuts` 请使用以下shortcuts.js：</span><span class="sxs-lookup"><span data-stu-id="55ecb-158">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="55ecb-159">属性名称 `action` ， `key` 和 `default` 是必需的。</span><span class="sxs-lookup"><span data-stu-id="55ecb-159">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="55ecb-160">该属性的值 `action` 是一个字符串，并且必须与操作对象 `id` 中的某个属性匹配。</span><span class="sxs-lookup"><span data-stu-id="55ecb-160">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="55ecb-161">该属性 `default` 可以是字符 A - Z、a -z、0 - 9 以及标点符号"-"、"_"和"+"的任意组合。</span><span class="sxs-lookup"><span data-stu-id="55ecb-161">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="55ecb-162"> (，这些属性中不使用小写字母。) </span><span class="sxs-lookup"><span data-stu-id="55ecb-162">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="55ecb-163">该属性必须包含至少一个修饰符键的名称 (`default` Alt、Ctrl、Shift) 一个其他键。</span><span class="sxs-lookup"><span data-stu-id="55ecb-163">The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.</span></span>
- <span data-ttu-id="55ecb-164">对于 Mac，我们还支持 COMMAND 修饰符键。</span><span class="sxs-lookup"><span data-stu-id="55ecb-164">For Macs, we also support the COMMAND modifier key.</span></span>
- <span data-ttu-id="55ecb-165">对于 Mac，ALT 映射到 OPTION 键。</span><span class="sxs-lookup"><span data-stu-id="55ecb-165">For Macs, ALT is mapped to the OPTION key.</span></span> <span data-ttu-id="55ecb-166">对于 Windows，COMMAND 映射到 Ctrl 键。</span><span class="sxs-lookup"><span data-stu-id="55ecb-166">For Windows, COMMAND is mapped to the CTRL key.</span></span>
- <span data-ttu-id="55ecb-167">当两个字符链接到标准键盘中的同一物理键时，它们是属性中的同义词;例如，Alt+a 和 Alt+A 是同一快捷方式 `default` ，Ctrl+- 和 Ctrl+ 也是，因为"-"和"_"是同一物理键。 \_</span><span class="sxs-lookup"><span data-stu-id="55ecb-167">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="55ecb-168">"+"字符指示同时按下其任一侧的键。</span><span class="sxs-lookup"><span data-stu-id="55ecb-168">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="55ecb-169">示例如下：</span><span class="sxs-lookup"><span data-stu-id="55ecb-169">The following is an example:</span></span>

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

<span data-ttu-id="55ecb-170">快捷方式 JSON 的完整架构位于extended-manifest.schema.js[on。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="55ecb-170">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="55ecb-171">Office 外接程序不支持键提示（也称为顺序键快捷方式，如 Excel 快捷方式选择填充颜色 **Alt+H、H）。**</span><span class="sxs-lookup"><span data-stu-id="55ecb-171">Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.</span></span>

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a><span data-ttu-id="55ecb-172">当焦点位于任务窗格中时，使用快捷方式</span><span class="sxs-lookup"><span data-stu-id="55ecb-172">Using shortcuts when the focus is in the task pane</span></span>

<span data-ttu-id="55ecb-173">目前，只有当用户的焦点位于工作表中时，才能调用 Office 外接程序的键盘快捷方式。</span><span class="sxs-lookup"><span data-stu-id="55ecb-173">Currently, the keyboard shortcuts for an Office Add-in can only be invoked when the user's focus is in the worksheet.</span></span> <span data-ttu-id="55ecb-174">当用户的焦点位于 Office UI (（如任务窗格) ）时，不会忽略任何加载项的快捷方式。</span><span class="sxs-lookup"><span data-stu-id="55ecb-174">When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored.</span></span> <span data-ttu-id="55ecb-175">作为一种解决方法，加载项可以定义键盘处理程序，当用户的焦点位于加载项 UI 内时，可以调用某些操作。</span><span class="sxs-lookup"><span data-stu-id="55ecb-175">As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.</span></span>

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a><span data-ttu-id="55ecb-176">使用 Office 或其他外接程序已使用的键组合</span><span class="sxs-lookup"><span data-stu-id="55ecb-176">Using key combinations that are already used by Office or another add-in</span></span>

<span data-ttu-id="55ecb-177">在预览期间，没有用于确定当用户按下由加载项以及 Office 或其他加载项注册的键组合时会发生什么情况的系统。</span><span class="sxs-lookup"><span data-stu-id="55ecb-177">During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in.</span></span> <span data-ttu-id="55ecb-178">行为未定义。</span><span class="sxs-lookup"><span data-stu-id="55ecb-178">Behavior is undefined.</span></span>

<span data-ttu-id="55ecb-179">目前，当两个或多个加载项已注册相同的键盘快捷方式时，没有解决方法，但您可以使用这些好的做法最大程度地减少与 Excel 的冲突：</span><span class="sxs-lookup"><span data-stu-id="55ecb-179">Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="55ecb-180">在外接程序中，仅使用以下模式的键盘快捷方式：\**Ctrl+Shift+Alt+* x\*\*\*，其中 *x* 是一些其他键。</span><span class="sxs-lookup"><span data-stu-id="55ecb-180">Use only keyboard shortcuts with the following pattern in your add-in: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="55ecb-181">如果您需要更多键盘快捷方式，请检查 [Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)键盘快捷方式的列表，并避免在外接程序中使用它们。</span><span class="sxs-lookup"><span data-stu-id="55ecb-181">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="55ecb-182">无法重写的浏览器快捷方式</span><span class="sxs-lookup"><span data-stu-id="55ecb-182">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="55ecb-183">不能使用下列任何键盘组合。</span><span class="sxs-lookup"><span data-stu-id="55ecb-183">You cannot use any of the following keyboard combinations.</span></span> <span data-ttu-id="55ecb-184">浏览器使用它们，并且不能重写。</span><span class="sxs-lookup"><span data-stu-id="55ecb-184">They are used by browsers and cannot be overridden.</span></span> <span data-ttu-id="55ecb-185">此列表是一项正在进行中的工作。</span><span class="sxs-lookup"><span data-stu-id="55ecb-185">This list is a work in progress.</span></span> <span data-ttu-id="55ecb-186">如果发现无法覆盖的其他组合，请使用此页面底部的反馈工具告诉我们。</span><span class="sxs-lookup"><span data-stu-id="55ecb-186">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="55ecb-187">Ctrl+N</span><span class="sxs-lookup"><span data-stu-id="55ecb-187">Ctrl+N</span></span>
- <span data-ttu-id="55ecb-188">Ctrl+Shift+N</span><span class="sxs-lookup"><span data-stu-id="55ecb-188">Ctrl+Shift+N</span></span>
- <span data-ttu-id="55ecb-189">Ctrl+T</span><span class="sxs-lookup"><span data-stu-id="55ecb-189">Ctrl+T</span></span>
- <span data-ttu-id="55ecb-190">Ctrl+Shift+T</span><span class="sxs-lookup"><span data-stu-id="55ecb-190">Ctrl+Shift+T</span></span>
- <span data-ttu-id="55ecb-191">Ctrl+W</span><span class="sxs-lookup"><span data-stu-id="55ecb-191">Ctrl+W</span></span>
- <span data-ttu-id="55ecb-192">Ctrl+PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="55ecb-192">Ctrl+PgUp/PgDn</span></span>

## <a name="localize-the-keyboard-shortcuts-json"></a><span data-ttu-id="55ecb-193">本地化键盘快捷方式 JSON</span><span class="sxs-lookup"><span data-stu-id="55ecb-193">Localize the keyboard shortcuts JSON</span></span>

<span data-ttu-id="55ecb-194">如果加载项支持多个区域设置，则需要本地化 `name` 操作对象的属性。</span><span class="sxs-lookup"><span data-stu-id="55ecb-194">If your add-in supports multiple locales, you'll need to localize the `name` property of the action objects.</span></span> <span data-ttu-id="55ecb-195">此外，如果外接程序支持的任何区域设置具有字母或不同的书写系统，因此使用不同的键盘，则您可能还需要本地化快捷方式。</span><span class="sxs-lookup"><span data-stu-id="55ecb-195">Also, if any of the locales that the add-in supports have alphabets or different writing systems, and hence different keyboards, you may need to localize the shortcuts also.</span></span> <span data-ttu-id="55ecb-196">若要了解如何本地化键盘快捷方式 JSON，请参阅 [本地化扩展替代](../develop/localization.md#localize-extended-overrides)。</span><span class="sxs-lookup"><span data-stu-id="55ecb-196">For information about how to localize the keyboard shortcuts JSON, see [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span></span>

## <a name="next-steps"></a><span data-ttu-id="55ecb-197">后续步骤</span><span class="sxs-lookup"><span data-stu-id="55ecb-197">Next Steps</span></span>

- <span data-ttu-id="55ecb-198">请参阅示例外接程序[excel-keyboard-shortcuts。](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="55ecb-198">See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>
- <span data-ttu-id="55ecb-199">获取在 Work 中处理扩展覆盖 [和清单的扩展覆盖的概述](../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="55ecb-199">Get an overview of working with extended overrides in [Work with extended overrides of the manifest](../develop/extended-overrides.md).</span></span>

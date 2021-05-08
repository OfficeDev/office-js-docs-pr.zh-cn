---
title: 加载项中的Office快捷方式
description: 了解如何将自定义键盘快捷方式（也称为组合键）Office加载项。
ms.date: 05/05/2021
localization_priority: Normal
ms.openlocfilehash: 42c0b5190d0fc71f137284950bcb983f16845fca
ms.sourcegitcommit: 132f5082f5bf9500dad0a2eaf89d924c823e575d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/07/2021
ms.locfileid: "52266105"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a><span data-ttu-id="04579-103">将自定义键盘快捷方式添加到Office加载项</span><span class="sxs-lookup"><span data-stu-id="04579-103">Add custom keyboard shortcuts to your Office Add-ins</span></span>

<span data-ttu-id="04579-104">键盘快捷方式（也称为组合键）使加载项的用户能够更高效地工作。</span><span class="sxs-lookup"><span data-stu-id="04579-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently.</span></span> <span data-ttu-id="04579-105">键盘快捷方式通过提供鼠标的替代方法，还可以为残障人士改进加载项的辅助功能。</span><span class="sxs-lookup"><span data-stu-id="04579-105">Keyboard shortcuts also improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="04579-106">若要从已启用键盘快捷方式的加载项的工作版本开始，请克隆并运行键盘快捷方式[Excel示例](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。</span><span class="sxs-lookup"><span data-stu-id="04579-106">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="04579-107">准备好向自己的加载项添加键盘快捷方式后，请继续阅读本文。</span><span class="sxs-lookup"><span data-stu-id="04579-107">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="04579-108">向加载项添加键盘快捷方式有三个步骤：</span><span class="sxs-lookup"><span data-stu-id="04579-108">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="04579-109">[配置加载项的清单](#configure-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="04579-109">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="04579-110">[创建或编辑快捷方式 JSON 文件](#create-or-edit-the-shortcuts-json-file) 以定义操作及其键盘快捷方式。</span><span class="sxs-lookup"><span data-stu-id="04579-110">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="04579-111">[添加](#create-a-mapping-of-actions-to-their-functions) [Office.actions.associate](/javascript/api/office/office.actions#associate) API 的一个或多个运行时调用，以将函数映射到每个操作。</span><span class="sxs-lookup"><span data-stu-id="04579-111">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="04579-112">配置清单</span><span class="sxs-lookup"><span data-stu-id="04579-112">Configure the manifest</span></span>

<span data-ttu-id="04579-113">清单有两个小更改需要进行。</span><span class="sxs-lookup"><span data-stu-id="04579-113">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="04579-114">一种是允许加载项使用共享运行时，另一种是指向定义键盘快捷方式的 JSON 格式文件。</span><span class="sxs-lookup"><span data-stu-id="04579-114">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="04579-115">将外接程序配置为使用共享运行时</span><span class="sxs-lookup"><span data-stu-id="04579-115">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="04579-116">添加自定义键盘快捷方式要求加载项使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="04579-116">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="04579-117">有关详细信息，请 [配置外接程序以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="04579-117">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="04579-118">将映射文件链接到清单</span><span class="sxs-lookup"><span data-stu-id="04579-118">Link the mapping file to the manifest</span></span>

<span data-ttu-id="04579-119">在 *紧* (不在) 元素内，添加 `<VersionOverrides>` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="04579-119">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="04579-120">将 `Url` 属性设置为项目中将在稍后步骤创建的 JSON 文件的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="04579-120">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="04579-121">创建或编辑快捷方式 JSON 文件</span><span class="sxs-lookup"><span data-stu-id="04579-121">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="04579-122">在项目中创建 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="04579-122">Create a JSON file in your project.</span></span> <span data-ttu-id="04579-123">确保文件的路径与为 `Url` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素的 属性指定的位置相匹配。</span><span class="sxs-lookup"><span data-stu-id="04579-123">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="04579-124">此文件将描述键盘快捷方式以及这些快捷方式将调用的操作。</span><span class="sxs-lookup"><span data-stu-id="04579-124">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="04579-125">在 JSON 文件中，有两个数组。</span><span class="sxs-lookup"><span data-stu-id="04579-125">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="04579-126">actions 数组将包含定义要调用的操作的对象，快捷方式数组将包含将键组合映射到操作的对象。</span><span class="sxs-lookup"><span data-stu-id="04579-126">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="04579-127">下面是一个示例：</span><span class="sxs-lookup"><span data-stu-id="04579-127">Here is an example:</span></span>

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

    <span data-ttu-id="04579-128">有关 JSON 对象详细信息，请参阅 [构造操作](#construct-the-action-objects) 对象和 [构造快捷方式对象](#construct-the-shortcut-objects)。</span><span class="sxs-lookup"><span data-stu-id="04579-128">For more information about the JSON objects, see [Construct the action objects](#construct-the-action-objects) and [Construct the shortcut objects](#construct-the-shortcut-objects).</span></span> <span data-ttu-id="04579-129">快捷方式 JSON 的完整架构位于 上的[extended-manifest.schema.js。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="04579-129">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="04579-130">可以在整个文章中使用"CONTROL"来表示"Ctrl"。</span><span class="sxs-lookup"><span data-stu-id="04579-130">You can use "CONTROL" in place of "Ctrl" throughout this article.</span></span>

    <span data-ttu-id="04579-131">在稍后的步骤中，操作本身将映射到您编写的函数。</span><span class="sxs-lookup"><span data-stu-id="04579-131">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="04579-132">此示例稍后将 SHOWTASKPANE 映射到调用 方法的函数， `Office.addin.showAsTaskpane` 将 HIDETASKPANE 映射到调用 该方法 `Office.addin.hide` 的函数。</span><span class="sxs-lookup"><span data-stu-id="04579-132">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="04579-133">创建操作到其函数的映射</span><span class="sxs-lookup"><span data-stu-id="04579-133">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="04579-134">在项目中，打开 元素中的 HTML 页面加载的 JavaScript `<FunctionFile>` 文件。</span><span class="sxs-lookup"><span data-stu-id="04579-134">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="04579-135">在 JavaScript 文件中，使用[Office.actions.associate](/javascript/api/office/office.actions#associate) API 将 JSON 文件中指定的每个操作映射到 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="04579-135">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="04579-136">将以下 JavaScript 添加到文件中。</span><span class="sxs-lookup"><span data-stu-id="04579-136">Add the following JavaScript to the file.</span></span> <span data-ttu-id="04579-137">关于代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="04579-137">Note the following about the code:</span></span>

    - <span data-ttu-id="04579-138">第一个参数是 JSON 文件的操作之一。</span><span class="sxs-lookup"><span data-stu-id="04579-138">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="04579-139">第二个参数是当用户按下映射到 JSON 文件中操作的组合键时运行的函数。</span><span class="sxs-lookup"><span data-stu-id="04579-139">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="04579-140">若要继续此示例，请使用 `'SHOWTASKPANE'` 作为第一个参数。</span><span class="sxs-lookup"><span data-stu-id="04579-140">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="04579-141">对于函数的正文，使用[Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--)方法打开加载项的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="04579-141">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="04579-142">完成后，代码应如下所示：</span><span class="sxs-lookup"><span data-stu-id="04579-142">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="04579-143">添加函数的第二个调用，以将操作映射到调用 `Office.actions.associate` `HIDETASKPANE` [Office.addin.hide 的函数](/javascript/api/office/office.addin#hide--)。</span><span class="sxs-lookup"><span data-stu-id="04579-143">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="04579-144">示例如下：</span><span class="sxs-lookup"><span data-stu-id="04579-144">The following is an example:</span></span>

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

<span data-ttu-id="04579-145">按照前面的步骤，加载项可通过按 **Ctrl+Alt+Up** 和 **Ctrl+Alt+Down 切换任务窗格的可见性**。</span><span class="sxs-lookup"><span data-stu-id="04579-145">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Alt+Up** and **Ctrl+Alt+Down**.</span></span> <span data-ttu-id="04579-146">相同的行为显示在 Excel 外接程序[](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)PnP Office中的键盘快捷方式示例GitHub。</span><span class="sxs-lookup"><span data-stu-id="04579-146">The same behavior is shown in the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample in the Office Add-ins PnP repo in GitHub.</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="04579-147">详细信息和限制</span><span class="sxs-lookup"><span data-stu-id="04579-147">Details and restrictions</span></span>

### <a name="construct-the-action-objects"></a><span data-ttu-id="04579-148">构造操作对象</span><span class="sxs-lookup"><span data-stu-id="04579-148">Construct the action objects</span></span>

<span data-ttu-id="04579-149">在 上指定对象数组中的对象时 `actions` ，shortcuts.js准则：</span><span class="sxs-lookup"><span data-stu-id="04579-149">Use the following guidelines when specifying the objects in the `actions` array of the shortcuts.json:</span></span>

- <span data-ttu-id="04579-150">属性名 `id` 和 `name` 是必需的。</span><span class="sxs-lookup"><span data-stu-id="04579-150">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="04579-151">`id`属性用于唯一标识使用键盘快捷方式调用的操作。</span><span class="sxs-lookup"><span data-stu-id="04579-151">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="04579-152">`name`属性必须是描述操作的用户友好字符串。</span><span class="sxs-lookup"><span data-stu-id="04579-152">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="04579-153">它必须是字符 A - Z、a - z、0 - 9 和标点符号"-"、"_"和"+"的组合。</span><span class="sxs-lookup"><span data-stu-id="04579-153">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="04579-154">属性是可选的。</span><span class="sxs-lookup"><span data-stu-id="04579-154">The `type` property is optional.</span></span> <span data-ttu-id="04579-155">当前仅 `ExecuteFunction` 支持类型。</span><span class="sxs-lookup"><span data-stu-id="04579-155">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="04579-156">示例如下：</span><span class="sxs-lookup"><span data-stu-id="04579-156">The following is an example:</span></span>

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

<span data-ttu-id="04579-157">快捷方式 JSON 的完整架构位于 上的[extended-manifest.schema.js。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="04579-157">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="construct-the-shortcut-objects"></a><span data-ttu-id="04579-158">构造快捷方式对象</span><span class="sxs-lookup"><span data-stu-id="04579-158">Construct the shortcut objects</span></span>

<span data-ttu-id="04579-159">在 上指定对象数组中的对象时 `shortcuts` ，shortcuts.js准则：</span><span class="sxs-lookup"><span data-stu-id="04579-159">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="04579-160">属性名称 `action` 、 `key` 和 `default` 是必需的。</span><span class="sxs-lookup"><span data-stu-id="04579-160">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="04579-161">该属性的值 `action` 是一个字符串，并且必须与 action 对象 `id` 中的某个属性匹配。</span><span class="sxs-lookup"><span data-stu-id="04579-161">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="04579-162">该属性 `default` 可以是字符 A - Z、-z、0 - 9 和标点符号"-"、"_"和"+"的任意组合。</span><span class="sxs-lookup"><span data-stu-id="04579-162">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="04579-163"> (根据惯例，这些属性中不使用小写字母。) </span><span class="sxs-lookup"><span data-stu-id="04579-163">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="04579-164">该属性 `default` 必须至少包含 Alt、Ctrl、Shift 和一个 (键的名称) 一个修饰符键。</span><span class="sxs-lookup"><span data-stu-id="04579-164">The `default` property must contain the name of at least one modifier key (Alt, Ctrl, Shift) and only one other key.</span></span>
- <span data-ttu-id="04579-165">对于 Mac，我们还支持 Command 修饰符键。</span><span class="sxs-lookup"><span data-stu-id="04579-165">For Macs, we also support the Command modifier key.</span></span>
- <span data-ttu-id="04579-166">对于 Mac，Alt 映射到 Option 键。</span><span class="sxs-lookup"><span data-stu-id="04579-166">For Macs, Alt is mapped to the Option key.</span></span> <span data-ttu-id="04579-167">例如Windows命令映射到 Ctrl 键。</span><span class="sxs-lookup"><span data-stu-id="04579-167">For Windows, Command is mapped to the Ctrl key.</span></span>
- <span data-ttu-id="04579-168">当两个字符链接到标准键盘中的同一个物理键时，它们是 属性中的同义词;例如，Alt+a 和 Alt+A 是同一快捷方式 `default` ，Ctrl+- 和 Ctrl+ 也是，因为 \_ "-"和"_"是同一个物理键。</span><span class="sxs-lookup"><span data-stu-id="04579-168">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, Alt+a and Alt+A are the same shortcut, so are Ctrl+- and Ctrl+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="04579-169">"+"字符指示同时按下其任一侧的键。</span><span class="sxs-lookup"><span data-stu-id="04579-169">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="04579-170">示例如下：</span><span class="sxs-lookup"><span data-stu-id="04579-170">The following is an example:</span></span>

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

<span data-ttu-id="04579-171">快捷方式 JSON 的完整架构位于 上的[extended-manifest.schema.js。](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="04579-171">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="04579-172">键提示（也称为连续键快捷方式，例如选择填充颜色的 Excel 快捷方式 **Alt+H、H）** 在加载项中不受Office支持。</span><span class="sxs-lookup"><span data-stu-id="04579-172">KeyTips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.</span></span>

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a><span data-ttu-id="04579-173">避免其他加载项使用组合键</span><span class="sxs-lookup"><span data-stu-id="04579-173">Avoid key combinations in use by other add-ins</span></span>

<span data-ttu-id="04579-174">有许多键盘快捷方式已由 Office。</span><span class="sxs-lookup"><span data-stu-id="04579-174">There are many keyboard shortcuts that are already in use by Office.</span></span> <span data-ttu-id="04579-175">避免为已在使用的外接程序注册键盘快捷方式，但在某些情况下，可能需要替代现有键盘快捷方式或处理已注册同一键盘快捷方式的多个加载项之间的冲突。</span><span class="sxs-lookup"><span data-stu-id="04579-175">Avoid registering keyboard shortcuts for your add-in that are already in use, however there may be some instances where it is necessary to override existing keyboard shortcuts or handle conflicts between multiple add-ins that have registered the same keyboard shortcut.</span></span>

<span data-ttu-id="04579-176">如果发生冲突，用户将在第一次尝试使用冲突的键盘快捷方式时看到一个对话框，请注意，此对话框中显示的动作名称是文件中 action 对象中的 属性。 `name` `shortcuts.json`</span><span class="sxs-lookup"><span data-stu-id="04579-176">In the case of a conflict, the user will see a dialog box the first time they attempt to use a conflicting keyboard shortcut, note that the action name that is displayed in this dialog is the `name` property in the action object in `shortcuts.json` file.</span></span>

![插图显示具有单个快捷方式的两个不同操作的冲突模式](../images/add-in-shortcut-conflict-modal.png)

<span data-ttu-id="04579-178">用户可以选择键盘快捷方式将执行的操作。</span><span class="sxs-lookup"><span data-stu-id="04579-178">The user can select which action the keyboard shortcut will take.</span></span> <span data-ttu-id="04579-179">做出选择后，保存首选项，供将来使用同一快捷方式。</span><span class="sxs-lookup"><span data-stu-id="04579-179">After making the selection, the preference is saved for future uses of the same shortcut.</span></span> <span data-ttu-id="04579-180">快捷方式首选项按用户、平台保存。</span><span class="sxs-lookup"><span data-stu-id="04579-180">The shortcut preferences are saved per user, per platform.</span></span> <span data-ttu-id="04579-181">如果用户希望更改其首选项，他们可以从"告诉我"搜索框中调用"重置Office外接程序快捷方式首选项"命令。 </span><span class="sxs-lookup"><span data-stu-id="04579-181">If the user wishes to change their preferences, they can invoke the **Reset Office Add-ins shortcut preferences** command from the **Tell me** search box.</span></span> <span data-ttu-id="04579-182">调用命令可清除用户的所有加载项快捷方式首选项，并且用户下次尝试使用冲突快捷方式时，会再次看到冲突对话框提示：</span><span class="sxs-lookup"><span data-stu-id="04579-182">Invoking the command clears all of the user's add-in shortcut preferences and the user will again be prompted with the conflict dialog box the next time they attempt to use a conflicting shortcut:</span></span>

![显示外接程序快捷方式首选项Excel重置Office中的"告诉我"搜索框](../images/add-in-reset-shortcuts-action.png)

<span data-ttu-id="04579-184">为了获得最佳用户体验，我们建议您尽量减少与以下Excel冲突：</span><span class="sxs-lookup"><span data-stu-id="04579-184">For the best user experience, we recommend that you minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="04579-185">请仅使用以下模式的键盘快捷方式： \**Ctrl+Shift+Alt+* x\*\*\*，其中 *x* 是一些其他键。</span><span class="sxs-lookup"><span data-stu-id="04579-185">Use only keyboard shortcuts with the following pattern: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="04579-186">如果您需要更多键盘快捷方式，请检查Excel[键盘](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)快捷方式的列表，并避免在外接程序中使用它们。</span><span class="sxs-lookup"><span data-stu-id="04579-186">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>
- <span data-ttu-id="04579-187">当键盘焦点位于加载项 UI 内时 **，Ctrl+空格** 键和 **Ctrl+Shift+F10** 将不起作用，因为这些都是基本的辅助功能快捷方式。</span><span class="sxs-lookup"><span data-stu-id="04579-187">When the keyboard focus is inside the add-in UI, **Ctrl+Spacebar** and **Ctrl+Shift+F10** will not work as these are essential accessibility shortcuts.</span></span>
- <span data-ttu-id="04579-188">在 Windows 或 Mac 计算机上，如果"重置 Office 外接程序快捷方式首选项"命令在搜索菜单上不可用，则用户可以通过通过上下文菜单自定义功能区，将该命令手动添加到功能区。</span><span class="sxs-lookup"><span data-stu-id="04579-188">On a Windows or Mac computer, if the "Reset Office Add-ins shortcut preferences" command is not available on the search menu, the user can manually add the command to the ribbon by customizing the ribbon through the context menu.</span></span>

## <a name="customize-the-keyboard-shortcuts-per-platform"></a><span data-ttu-id="04579-189">自定义每个平台的键盘快捷方式</span><span class="sxs-lookup"><span data-stu-id="04579-189">Customize the keyboard shortcuts per platform</span></span>

<span data-ttu-id="04579-190">可以自定义特定于平台的快捷方式。</span><span class="sxs-lookup"><span data-stu-id="04579-190">It's possible to customize shortcuts to be platform-specific.</span></span> <span data-ttu-id="04579-191">下面是自定义以下每个平台的快捷方式的对象示例 `shortcuts` `windows` `mac` ：、、。 `web`</span><span class="sxs-lookup"><span data-stu-id="04579-191">The following is an example of the `shortcuts` object that customizes the shortcuts for each of the following platforms: `windows`, `mac`, `web`.</span></span> <span data-ttu-id="04579-192">请注意，您仍必须具有 `default` 每个快捷方式的快捷键。</span><span class="sxs-lookup"><span data-stu-id="04579-192">Note that you must still have a `default` shortcut key for each shortcut.</span></span>

<span data-ttu-id="04579-193">在下面的示例中， `default` 键是未指定的任何平台的回退键。</span><span class="sxs-lookup"><span data-stu-id="04579-193">In the following example, the `default` key is the fallback key for any platform that is not specified.</span></span> <span data-ttu-id="04579-194">唯一未指定的平台Windows，因此 `default` 该密钥仅适用于Windows。</span><span class="sxs-lookup"><span data-stu-id="04579-194">The only platform not specified is Windows, so the `default` key will only apply to Windows.</span></span>

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

## <a name="localize-the-keyboard-shortcuts-json"></a><span data-ttu-id="04579-195">本地化键盘快捷方式 JSON</span><span class="sxs-lookup"><span data-stu-id="04579-195">Localize the keyboard shortcuts JSON</span></span>

<span data-ttu-id="04579-196">如果加载项支持多个区域设置，则需要本地化 `name` action 对象的 属性。</span><span class="sxs-lookup"><span data-stu-id="04579-196">If your add-in supports multiple locales, you'll need to localize the `name` property of the action objects.</span></span> <span data-ttu-id="04579-197">此外，如果加载项支持的任何区域设置具有字母或不同的书写系统，因此使用不同的键盘，则你可能还需要本地化快捷方式。</span><span class="sxs-lookup"><span data-stu-id="04579-197">Also, if any of the locales that the add-in supports have alphabets or different writing systems, and hence different keyboards, you may need to localize the shortcuts also.</span></span> <span data-ttu-id="04579-198">若要了解如何本地化键盘快捷方式 JSON，请参阅 [本地化扩展替代](../develop/localization.md#localize-extended-overrides)。</span><span class="sxs-lookup"><span data-stu-id="04579-198">For information about how to localize the keyboard shortcuts JSON, see [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="04579-199">无法重写的浏览器快捷方式</span><span class="sxs-lookup"><span data-stu-id="04579-199">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="04579-200">在 Web 上使用自定义键盘快捷方式时，外接程序无法覆盖浏览器所使用的某些键盘快捷方式。此列表是一项正在进行中的工作。</span><span class="sxs-lookup"><span data-stu-id="04579-200">When using custom keyboard shortcuts on the web, some keyboard shortcuts that are used by the browser cannot be overridden by add-ins. This list is a work in progress.</span></span> <span data-ttu-id="04579-201">如果发现无法覆盖的其他组合，请使用此页面底部的反馈工具告诉我们。</span><span class="sxs-lookup"><span data-stu-id="04579-201">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="04579-202">Ctrl+N</span><span class="sxs-lookup"><span data-stu-id="04579-202">Ctrl+N</span></span>
- <span data-ttu-id="04579-203">Ctrl+Shift+N</span><span class="sxs-lookup"><span data-stu-id="04579-203">Ctrl+Shift+N</span></span>
- <span data-ttu-id="04579-204">Ctrl+T</span><span class="sxs-lookup"><span data-stu-id="04579-204">Ctrl+T</span></span>
- <span data-ttu-id="04579-205">Ctrl+Shift+T</span><span class="sxs-lookup"><span data-stu-id="04579-205">Ctrl+Shift+T</span></span>
- <span data-ttu-id="04579-206">Ctrl+W</span><span class="sxs-lookup"><span data-stu-id="04579-206">Ctrl+W</span></span>
- <span data-ttu-id="04579-207">Ctrl+PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="04579-207">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="04579-208">后续步骤</span><span class="sxs-lookup"><span data-stu-id="04579-208">Next Steps</span></span>

- <span data-ttu-id="04579-209">请参阅[Excel键盘快捷方式](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)示例外接程序。</span><span class="sxs-lookup"><span data-stu-id="04579-209">See the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample add-in.</span></span>
- <span data-ttu-id="04579-210">获取有关使用清单的扩展替代 [中的扩展覆盖的概述](../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="04579-210">Get an overview of working with extended overrides in [Work with extended overrides of the manifest](../develop/extended-overrides.md).</span></span>

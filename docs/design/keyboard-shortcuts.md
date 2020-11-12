---
title: Office 外接程序中的自定义键盘快捷方式
description: 了解如何将自定义键盘快捷方式（也称为键组合）添加到 Office 外接程序。
ms.date: 11/09/2020
localization_priority: Normal
ms.openlocfilehash: f95c26067203a4ec2659aa6a632403c96ed81674
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996677"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a><span data-ttu-id="bbfd8-103">将自定义键盘快捷方式添加到 Office 外接 (预览) </span><span class="sxs-lookup"><span data-stu-id="bbfd8-103">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>

<span data-ttu-id="bbfd8-104">键盘快捷方式（也称为键组合）使您的外接程序的用户可以更高效地工作，并通过提供鼠标替换功能为残障用户改进了加载项的辅助功能。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="bbfd8-105">若要从已启用的键盘快捷方式开始使用加载项的工作版本，请克隆并运行示例 [Excel 键盘快捷方式](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-105">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="bbfd8-106">准备好将键盘快捷方式添加到自己的外接程序后，请继续阅读本文。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-106">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="bbfd8-107">将键盘快捷方式添加到外接程序中有三个步骤：</span><span class="sxs-lookup"><span data-stu-id="bbfd8-107">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="bbfd8-108">[配置加载项的清单](#configure-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-108">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="bbfd8-109">[创建或编辑快捷方式 JSON 文件](#create-or-edit-the-shortcuts-json-file) 以定义操作及其键盘快捷方式。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-109">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="bbfd8-110">[添加一个或多个 Office 的运行时调用](#create-a-mapping-of-actions-to-their-functions) [。关联](/javascript/api/office/office.actions#associate) API 以将某个函数映射到每个操作。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-110">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="bbfd8-111">配置清单</span><span class="sxs-lookup"><span data-stu-id="bbfd8-111">Configure the manifest</span></span>

<span data-ttu-id="bbfd8-112">对清单进行了两处较小的更改。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-112">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="bbfd8-113">一种是使外接程序能够使用共享运行时，而另一种是指向您定义了键盘快捷方式的 JSON 格式的文件。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-113">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="bbfd8-114">将加载项配置为使用共享运行时</span><span class="sxs-lookup"><span data-stu-id="bbfd8-114">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="bbfd8-115">若要添加自定义键盘快捷方式，您的加载项需要使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-115">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="bbfd8-116">有关详细信息，请 [配置外接程序以使用共享运行时](../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-116">For more information, [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="bbfd8-117">将映射文件链接到清单</span><span class="sxs-lookup"><span data-stu-id="bbfd8-117">Link the mapping file to the manifest</span></span>

<span data-ttu-id="bbfd8-118">在 *下面* 紧接着 (不在 `<VersionOverrides>` 清单中的元素) 元素中，添加一个 [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-118">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="bbfd8-119">将 `Url` 属性设置为项目中您将在后续步骤中创建的 JSON 文件的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-119">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="bbfd8-120">创建或编辑快捷方式 JSON 文件</span><span class="sxs-lookup"><span data-stu-id="bbfd8-120">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="bbfd8-121">在项目中创建一个 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-121">Create a JSON file in your project.</span></span> <span data-ttu-id="bbfd8-122">确保文件的路径与您为 ExtendedOverrides 元素的属性指定的位置相匹配 `Url` 。 [ExtendedOverrides](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="bbfd8-122">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="bbfd8-123">此文件将介绍你的键盘快捷方式以及它们将调用的操作。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-123">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="bbfd8-124">在 JSON 文件中，有两个数组。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-124">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="bbfd8-125">操作数组将包含定义要调用的操作的对象，并且快捷键数组将包含将键组合映射到操作的对象。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-125">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="bbfd8-126">如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="bbfd8-126">Here is an example:</span></span>

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

    <span data-ttu-id="bbfd8-127">有关 JSON 对象的详细信息，请参阅 [构造 action 对象](#constructing-the-action-objects) 和 [构造快捷方式对象](#constructing-the-shortcut-objects)。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-127">For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span></span> <span data-ttu-id="bbfd8-128">快捷键 JSON 的完整架构位于 [extended-manifest.schema.js](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="bbfd8-129">在本文中，您可以使用 "控制" 代替 "CTRL"。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-129">You can use "CONTROL" in place of "CTRL" throughout this article.</span></span>

    <span data-ttu-id="bbfd8-130">在后续步骤中，操作本身将映射到您编写的函数。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-130">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="bbfd8-131">在此示例中，您稍后会将 SHOWTASKPANE 映射到一个函数，该函数调用 `Office.addin.showAsTaskpane` 方法和 HIDETASKPANE 到调用该 `Office.addin.hide` 方法的函数。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-131">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="bbfd8-132">创建操作到它们的函数的映射</span><span class="sxs-lookup"><span data-stu-id="bbfd8-132">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="bbfd8-133">在您的项目中，打开元素中的 HTML 页面加载的 JavaScript 文件 `<FunctionFile>` 。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-133">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="bbfd8-134">在 JavaScript 文件中，使用 [Office. 操作。关联](/javascript/api/office/office.actions#associate) API 将您在 JSON 文件中指定的每个操作映射到一个 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-134">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="bbfd8-135">向文件中添加以下 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-135">Add the following JavaScript to the file.</span></span> <span data-ttu-id="bbfd8-136">请注意有关代码的以下内容：</span><span class="sxs-lookup"><span data-stu-id="bbfd8-136">Note the following about the code:</span></span>

    - <span data-ttu-id="bbfd8-137">第一个参数是 JSON 文件中的一项操作。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-137">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="bbfd8-138">第二个参数是当用户按下将映射到 JSON 文件中的操作的组合键时运行的函数。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-138">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="bbfd8-139">若要继续本示例，请使用 `'SHOWTASKPANE'` 作为第一个参数。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-139">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="bbfd8-140">对于函数的主体，请使用 [showTaskpane](/javascript/api/office/office.addin.md#showastaskpane--) 方法打开外接程序的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-140">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin.md#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="bbfd8-141">完成后，代码应类似于以下内容：</span><span class="sxs-lookup"><span data-stu-id="bbfd8-141">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="bbfd8-142">添加第二个函数调用， `Office.actions.associate` 以将 `HIDETASKPANE` 操作映射到一个调用了 [.addin](/javascript/api/office/office.addin.md#hide--)的函数。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-142">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin.md#hide--).</span></span> <span data-ttu-id="bbfd8-143">示例如下：</span><span class="sxs-lookup"><span data-stu-id="bbfd8-143">The following is an example:</span></span>

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

<span data-ttu-id="bbfd8-144">按照前面的步骤，你的外接程序可以通过按 **ctrl + shift + 向上箭头键** 和 **Ctrl + Shift + 向下箭头键** 来切换任务窗格的可见性。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-144">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**.</span></span> <span data-ttu-id="bbfd8-145">这与 [示例 excel 键盘快捷方式加载项](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)中所示的行为相同。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-145">This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="bbfd8-146">详细信息和限制</span><span class="sxs-lookup"><span data-stu-id="bbfd8-146">Details and restrictions</span></span>

### <a name="constructing-the-action-objects"></a><span data-ttu-id="bbfd8-147">构造 action 对象</span><span class="sxs-lookup"><span data-stu-id="bbfd8-147">Constructing the action objects</span></span>

<span data-ttu-id="bbfd8-148">在中指定 shortcuts.js数组中的对象时，请使用以下准则 `action` ：</span><span class="sxs-lookup"><span data-stu-id="bbfd8-148">Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:</span></span>

- <span data-ttu-id="bbfd8-149">属性名称 `id` ，并且 `name` 是强制性的。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-149">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="bbfd8-150">该 `id` 属性用于唯一标识要使用键盘快捷方式调用的操作。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-150">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="bbfd8-151">该 `name` 属性必须是描述操作的用户友好字符串。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-151">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="bbfd8-152">它必须是字符 a-z、a-z、0-9 和标点符号 "-"、"_" 和 "+" 的组合。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-152">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="bbfd8-153">属性是可选的。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-153">The `type` property is optional.</span></span> <span data-ttu-id="bbfd8-154">目前仅 `ExecuteFunction` 支持类型。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-154">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="bbfd8-155">示例如下：</span><span class="sxs-lookup"><span data-stu-id="bbfd8-155">The following is an example:</span></span>

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

<span data-ttu-id="bbfd8-156">快捷键 JSON 的完整架构位于 [extended-manifest.schema.js](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-156">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="constructing-the-shortcut-objects"></a><span data-ttu-id="bbfd8-157">构造快捷方式对象</span><span class="sxs-lookup"><span data-stu-id="bbfd8-157">Constructing the shortcut objects</span></span>

<span data-ttu-id="bbfd8-158">在中指定 shortcuts.js数组中的对象时，请使用以下准则 `shortcuts` ：</span><span class="sxs-lookup"><span data-stu-id="bbfd8-158">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="bbfd8-159">属性名称 `action` 、 `key` 和 `default` 是必需的。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-159">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="bbfd8-160">该属性的值 `action` 是一个字符串，并且必须与 `id` action 对象中的一个属性相匹配。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-160">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="bbfd8-161">该 `default` 属性可以是字符 a-z、a-z、0-9 和标点符号 "-"、"_" 和 "+" 的任意组合。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-161">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="bbfd8-162"> (按惯例，在这些属性中不使用小写字母。 ) </span><span class="sxs-lookup"><span data-stu-id="bbfd8-162">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="bbfd8-163">`default`属性必须包含至少一个修改键的名称 (ALT、CTRL、SHIFT) 且仅包含一个其他键。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-163">The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.</span></span>
- <span data-ttu-id="bbfd8-164">对于 Mac，我们还支持命令修改键。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-164">For Macs, we also support the COMMAND modifier key.</span></span>
- <span data-ttu-id="bbfd8-165">对于 Mac，将 ALT 映射到选项键。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-165">For Macs, ALT is mapped to the OPTION key.</span></span> <span data-ttu-id="bbfd8-166">对于 Windows，命令映射到 CTRL 键。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-166">For Windows, COMMAND is mapped to the CTRL key.</span></span>
- <span data-ttu-id="bbfd8-167">当两个字符链接到标准键盘中的同一个物理键时，它们就是属性中的同义词 `default` ; 例如，alt + a 和 alt + a 是相同的快捷方式，因此是 ctrl +-和 ctrl +， \_ 因为 "-" 和 "_" 是相同的物理键。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-167">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="bbfd8-168">"+" 字符指示同时按下的键的任意一侧的键。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-168">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="bbfd8-169">示例如下：</span><span class="sxs-lookup"><span data-stu-id="bbfd8-169">The following is an example:</span></span>

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

<span data-ttu-id="bbfd8-170">快捷键 JSON 的完整架构位于 [extended-manifest.schema.js](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-170">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="bbfd8-171">快捷键提示（也称为连续键快捷方式，例如，用于选择填充颜色的 Excel 快捷方式 **Alt + h，h** ）在 Office 加载项中不受支持。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-171">Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H** , are not supported in Office add-ins.</span></span>

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a><span data-ttu-id="bbfd8-172">当焦点在任务窗格中时使用快捷方式</span><span class="sxs-lookup"><span data-stu-id="bbfd8-172">Using shortcuts when the focus is in the task pane</span></span>

<span data-ttu-id="bbfd8-173">目前，只有当用户的焦点在工作表中时，才能调用 Office 外接程序的键盘快捷方式。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-173">Currently, the keyboard shortcuts for an Office add-in can only be invoked when the user's focus is in the worksheet.</span></span> <span data-ttu-id="bbfd8-174">当用户的焦点位于 Office UI (（例如任务窗格) ）中时，不会忽略任何加载项的快捷方式。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-174">When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored.</span></span> <span data-ttu-id="bbfd8-175">作为一种解决方法，加载项可以定义键盘处理程序，当用户的焦点位于外接程序 UI 中时，可以调用某些操作。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-175">As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.</span></span>

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a><span data-ttu-id="bbfd8-176">使用已由 Office 或其他加载项使用的组合键</span><span class="sxs-lookup"><span data-stu-id="bbfd8-176">Using key combinations that are already used by Office or another add-in</span></span>

<span data-ttu-id="bbfd8-177">在预览期间，没有系统可用于确定当用户按外接程序注册的组合键以及由 Office 或其他外接程序注册时，会发生什么情况。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-177">During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in.</span></span> <span data-ttu-id="bbfd8-178">行为未定义。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-178">Behavior is undefined.</span></span>

<span data-ttu-id="bbfd8-179">目前，如果两个或更多个加载项注册了相同的键盘快捷方式，但您可以最大限度地减少与 Excel 的冲突，请使用以下这些好的做法：</span><span class="sxs-lookup"><span data-stu-id="bbfd8-179">Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="bbfd8-180">在外接程序中仅使用具有以下模式的键盘快捷方式： \* *Ctrl + Shift + Alt +* x \* \* \*，其中 *x* 是另一个键。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-180">Use only keyboard shortcuts with the following pattern in your add-in: \* *Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="bbfd8-181">如果需要更多键盘快捷方式，请查看 [Excel 键盘快捷方式列表](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)，并避免在外接程序中使用其中任何一个。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-181">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="bbfd8-182">无法覆盖的浏览器快捷方式</span><span class="sxs-lookup"><span data-stu-id="bbfd8-182">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="bbfd8-183">您不能使用以下任何键盘组合。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-183">You cannot use any of the following keyboard combinations.</span></span> <span data-ttu-id="bbfd8-184">它们由浏览器使用，不能覆盖。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-184">They are used by browsers and cannot be overridden.</span></span> <span data-ttu-id="bbfd8-185">此列表是一项正在进行的工作。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-185">This list is a work in progress.</span></span> <span data-ttu-id="bbfd8-186">如果发现无法覆盖的其他组合，请使用本页底部的反馈工具告知我们。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-186">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="bbfd8-187">Ctrl + N</span><span class="sxs-lookup"><span data-stu-id="bbfd8-187">Ctrl+N</span></span>
- <span data-ttu-id="bbfd8-188">Ctrl + Shift + N</span><span class="sxs-lookup"><span data-stu-id="bbfd8-188">Ctrl+Shift+N</span></span>
- <span data-ttu-id="bbfd8-189">Ctrl + T</span><span class="sxs-lookup"><span data-stu-id="bbfd8-189">Ctrl+T</span></span>
- <span data-ttu-id="bbfd8-190">Ctrl + Shift + T</span><span class="sxs-lookup"><span data-stu-id="bbfd8-190">Ctrl+Shift+T</span></span>
- <span data-ttu-id="bbfd8-191">Ctrl + W</span><span class="sxs-lookup"><span data-stu-id="bbfd8-191">Ctrl+W</span></span>
- <span data-ttu-id="bbfd8-192">Ctrl + PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="bbfd8-192">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="bbfd8-193">后续步骤</span><span class="sxs-lookup"><span data-stu-id="bbfd8-193">Next Steps</span></span>

- <span data-ttu-id="bbfd8-194">请参阅示例加载项 [excel-键盘快捷方式](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。</span><span class="sxs-lookup"><span data-stu-id="bbfd8-194">See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

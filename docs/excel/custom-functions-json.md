---
ms.date: 01/14/2020
description: 在 Excel 中定义自定义函数的 JSON 元数据，并将您的函数 id 和 name 属性相关联。
title: Excel 中自定义函数的元数据
localization_priority: Normal
ms.openlocfilehash: 2a777cb0217d48caf03983d3dbfe662dfe0b2567
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217030"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="96f3d-103">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="96f3d-103">Custom functions metadata</span></span>

<span data-ttu-id="96f3d-104">如 "[自定义函数概述](custom-functions-overview.md)" 一文中所述，自定义函数项目必须包括 JSON 元数据文件和脚本（JavaScript 或 TypeScript）文件才能注册函数，使其可供使用。</span><span class="sxs-lookup"><span data-stu-id="96f3d-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="96f3d-105">自定义函数在用户首次运行外接程序且在所有工作簿中对同一用户可用时注册。</span><span class="sxs-lookup"><span data-stu-id="96f3d-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="96f3d-106">建议您尽可能使用 JSON 自动生成，方法是使用`yo office`搭建文件，这与[Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)中所示的过程类似，这是因为此过程更简单且不易出现用户错误。</span><span class="sxs-lookup"><span data-stu-id="96f3d-106">It is recommended that you use JSON autogeneration when possible, using the `yo office` scaffold files, similar to the process shown in the [Excel Custom Function tutorial](../tutorials/excel-tutorial-create-custom-functions.md) because this process is easier and less prone to user error.</span></span> <span data-ttu-id="96f3d-107">有关 JSDoc 注释 JSON 文件生成的过程的详细信息，请参阅[为自定义函数生成 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="96f3d-107">For more information on the process of JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="96f3d-108">不过，您可以从头开始创建自定义函数项目;它要求您执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="96f3d-108">However, you can make a custom functions project from scratch; it requires that you:</span></span>

- <span data-ttu-id="96f3d-109">手动编写 JSON 文件</span><span class="sxs-lookup"><span data-stu-id="96f3d-109">Write your JSON file by hand</span></span>
- <span data-ttu-id="96f3d-110">检查您的清单文件是否已连接到手动创作的 JSON 文件</span><span class="sxs-lookup"><span data-stu-id="96f3d-110">Check that your manifest file is connected to your hand-authored JSON file</span></span>
- <span data-ttu-id="96f3d-111">在脚本文件中`id`关联`name`函数和属性，以便注册您的函数</span><span class="sxs-lookup"><span data-stu-id="96f3d-111">Associate your functions' `id` and `name` properties in the script file in order to register your functions</span></span>

<span data-ttu-id="96f3d-112">本文将介绍如何执行所有三个步骤。</span><span class="sxs-lookup"><span data-stu-id="96f3d-112">This article will show you how to do all three of these steps.</span></span>

<span data-ttu-id="96f3d-113">下图说明了使用`yo office`搭建文件和从草稿写入 JSON 之间的差异。</span><span class="sxs-lookup"><span data-stu-id="96f3d-113">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>
<span data-ttu-id="96f3d-114">![使用 Yo 办公室和编写自己的 JSON 的差异的图像](../images/custom-functions-json.png)</span><span class="sxs-lookup"><span data-stu-id="96f3d-114">![Image of differences between using Yo Office and writing your own JSON](../images/custom-functions-json.png)</span></span>

> [!NOTE]
> <span data-ttu-id="96f3d-115">与`yo office`搭建文件相比，您需要通过 XML 清单文件中的`<Resources>`节将清单连接到所创建的 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="96f3d-115">In contrast with the `yo office` scaffold files, you need to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file.</span></span> <span data-ttu-id="96f3d-116">请注意，承载 JSON 文件的服务器上的服务器设置必须启用了[CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) ，才能使自定义函数在 web 上的 Excel 中正常工作。</span><span class="sxs-lookup"><span data-stu-id="96f3d-116">Note that the server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="96f3d-117">创作元数据并连接到清单</span><span class="sxs-lookup"><span data-stu-id="96f3d-117">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="96f3d-118">您需要在项目中创建一个 JSON 文件，并提供有关函数的所有详细信息，如函数的参数。</span><span class="sxs-lookup"><span data-stu-id="96f3d-118">You need to create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="96f3d-119">有关函数属性的完整列表，请参阅[以下元数据示例](#json-metadata-example)和[元数据参考](#metadata-reference)。</span><span class="sxs-lookup"><span data-stu-id="96f3d-119">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="96f3d-120">您还需要确保您的 XML 清单文件引用您在部分中的`<Resources>` JSON 文件，类似于以下示例。</span><span class="sxs-lookup"><span data-stu-id="96f3d-120">You also need to make sure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a><span data-ttu-id="96f3d-121">JSON 元数据示例</span><span class="sxs-lookup"><span data-stu-id="96f3d-121">JSON metadata example</span></span>

<span data-ttu-id="96f3d-122">以下示例介绍了定义自定义函数的加载项的 JSON 元数据文件的内容。</span><span class="sxs-lookup"><span data-stu-id="96f3d-122">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="96f3d-123">此示例后面的部分提供了有关此 JSON 示例中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="96f3d-123">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

```json
{
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "first",
          "description": "first number to add",
          "type": "number",
          "dimensionality": "scalar"
        },
        {
          "name": "second",
          "description": "second number to add",
          "type": "number",
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE",
      "description": "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "increment",
          "description": "the number to be added each time",
          "type": "number",
          "dimensionality": "scalar"
        }
      ],
      "options": {
        "stream": true,
        "cancelable": true
      }
    },
    {
      "id": "SECONDHIGHEST",
      "name": "SECONDHIGHEST",
      "description": "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "range",
          "description": "the input range",
          "type": "number",
          "dimensionality": "matrix"
        }
      ]
    }
  ]
}
```

> [!NOTE]
> <span data-ttu-id="96f3d-124">[OfficeDev/Excel 自定义函数](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json)GitHub 存储库的提交历史记录中提供了完整的示例 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="96f3d-124">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="96f3d-125">随着项目已调整为自动生成 JSON，手写 JSON 的完整示例仅在项目的早期版本中可用。</span><span class="sxs-lookup"><span data-stu-id="96f3d-125">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="96f3d-126">元数据参考</span><span class="sxs-lookup"><span data-stu-id="96f3d-126">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="96f3d-127">functions</span><span class="sxs-lookup"><span data-stu-id="96f3d-127">functions</span></span>

<span data-ttu-id="96f3d-128">`functions` 属性是自定义函数对象的一个数组。</span><span class="sxs-lookup"><span data-stu-id="96f3d-128">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="96f3d-129">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="96f3d-129">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="96f3d-130">属性</span><span class="sxs-lookup"><span data-stu-id="96f3d-130">Property</span></span>      | <span data-ttu-id="96f3d-131">数据类型</span><span class="sxs-lookup"><span data-stu-id="96f3d-131">Data type</span></span> | <span data-ttu-id="96f3d-132">必需</span><span class="sxs-lookup"><span data-stu-id="96f3d-132">Required</span></span> | <span data-ttu-id="96f3d-133">说明</span><span class="sxs-lookup"><span data-stu-id="96f3d-133">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="96f3d-134">string</span><span class="sxs-lookup"><span data-stu-id="96f3d-134">string</span></span>    | <span data-ttu-id="96f3d-135">No</span><span class="sxs-lookup"><span data-stu-id="96f3d-135">No</span></span>       | <span data-ttu-id="96f3d-136">最终用户在 Excel 中看到的函数的说明。</span><span class="sxs-lookup"><span data-stu-id="96f3d-136">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="96f3d-137">例如，**将摄氏度值转换为华氏度**。</span><span class="sxs-lookup"><span data-stu-id="96f3d-137">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="96f3d-138">string</span><span class="sxs-lookup"><span data-stu-id="96f3d-138">string</span></span>    | <span data-ttu-id="96f3d-139">No</span><span class="sxs-lookup"><span data-stu-id="96f3d-139">No</span></span>       | <span data-ttu-id="96f3d-140">提供有关函数的信息的 URL。</span><span class="sxs-lookup"><span data-stu-id="96f3d-140">URL that provides information about the function.</span></span> <span data-ttu-id="96f3d-141">（它显示在任务窗格中。）例如，`http://contoso.com/help/convertcelsiustofahrenheit.html`。</span><span class="sxs-lookup"><span data-stu-id="96f3d-141">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="96f3d-142">string</span><span class="sxs-lookup"><span data-stu-id="96f3d-142">string</span></span>    | <span data-ttu-id="96f3d-143">是</span><span class="sxs-lookup"><span data-stu-id="96f3d-143">Yes</span></span>      | <span data-ttu-id="96f3d-144">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="96f3d-144">A unique ID for the function.</span></span> <span data-ttu-id="96f3d-145">此 ID 只能包含字母数字字符和句点，设置后不应更改。</span><span class="sxs-lookup"><span data-stu-id="96f3d-145">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="96f3d-146">string</span><span class="sxs-lookup"><span data-stu-id="96f3d-146">string</span></span>    | <span data-ttu-id="96f3d-147">是</span><span class="sxs-lookup"><span data-stu-id="96f3d-147">Yes</span></span>      | <span data-ttu-id="96f3d-148">最终用户在 Excel 中看到的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="96f3d-148">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="96f3d-149">在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。</span><span class="sxs-lookup"><span data-stu-id="96f3d-149">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="96f3d-150">object</span><span class="sxs-lookup"><span data-stu-id="96f3d-150">object</span></span>    | <span data-ttu-id="96f3d-151">No</span><span class="sxs-lookup"><span data-stu-id="96f3d-151">No</span></span>       | <span data-ttu-id="96f3d-152">使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="96f3d-152">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="96f3d-153">有关详细信息，请参阅[选项](#options)。</span><span class="sxs-lookup"><span data-stu-id="96f3d-153">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="96f3d-154">array</span><span class="sxs-lookup"><span data-stu-id="96f3d-154">array</span></span>     | <span data-ttu-id="96f3d-155">是</span><span class="sxs-lookup"><span data-stu-id="96f3d-155">Yes</span></span>      | <span data-ttu-id="96f3d-156">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="96f3d-156">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="96f3d-157">有关详细信息，请参阅[参数](#parameters)。</span><span class="sxs-lookup"><span data-stu-id="96f3d-157">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="96f3d-158">object</span><span class="sxs-lookup"><span data-stu-id="96f3d-158">object</span></span>    | <span data-ttu-id="96f3d-159">是</span><span class="sxs-lookup"><span data-stu-id="96f3d-159">Yes</span></span>      | <span data-ttu-id="96f3d-160">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="96f3d-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="96f3d-161">有关详细信息，请参阅[结果](#result)。</span><span class="sxs-lookup"><span data-stu-id="96f3d-161">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="96f3d-162">options</span><span class="sxs-lookup"><span data-stu-id="96f3d-162">options</span></span>

<span data-ttu-id="96f3d-163">`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="96f3d-163">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="96f3d-164">下表列出了 `options` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="96f3d-164">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="96f3d-165">属性</span><span class="sxs-lookup"><span data-stu-id="96f3d-165">Property</span></span>          | <span data-ttu-id="96f3d-166">数据类型</span><span class="sxs-lookup"><span data-stu-id="96f3d-166">Data type</span></span> | <span data-ttu-id="96f3d-167">必需</span><span class="sxs-lookup"><span data-stu-id="96f3d-167">Required</span></span>                               | <span data-ttu-id="96f3d-168">说明</span><span class="sxs-lookup"><span data-stu-id="96f3d-168">Description</span></span>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | <span data-ttu-id="96f3d-169">boolean</span><span class="sxs-lookup"><span data-stu-id="96f3d-169">boolean</span></span>   | <span data-ttu-id="96f3d-170">否</span><span class="sxs-lookup"><span data-stu-id="96f3d-170">No</span></span><br/><br/><span data-ttu-id="96f3d-171">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="96f3d-171">Default value is `false`.</span></span>  | <span data-ttu-id="96f3d-172">如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `CancelableInvocation` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。</span><span class="sxs-lookup"><span data-stu-id="96f3d-172">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="96f3d-173">可取消函数通常仅用于返回单个结果的异步函数，并需要处理对数据请求的取消操作。</span><span class="sxs-lookup"><span data-stu-id="96f3d-173">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="96f3d-174">函数不能同时为流式处理和可取消。</span><span class="sxs-lookup"><span data-stu-id="96f3d-174">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="96f3d-175">有关详细信息，请参阅[Make a 流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)结尾附近的注释。</span><span class="sxs-lookup"><span data-stu-id="96f3d-175">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="96f3d-176">boolean</span><span class="sxs-lookup"><span data-stu-id="96f3d-176">boolean</span></span>   | <span data-ttu-id="96f3d-177">否</span><span class="sxs-lookup"><span data-stu-id="96f3d-177">No</span></span> <br/><br/><span data-ttu-id="96f3d-178">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="96f3d-178">Default value is `false`.</span></span> | <span data-ttu-id="96f3d-179">如果`true`为，则自定义函数可以访问调用自定义函数的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="96f3d-179">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="96f3d-180">若要获取调用自定义函数的单元格的地址，请在自定义函数中使用 context。</span><span class="sxs-lookup"><span data-stu-id="96f3d-180">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="96f3d-181">有关详细信息，请参阅[寻址单元格的上下文参数](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter)。</span><span class="sxs-lookup"><span data-stu-id="96f3d-181">For more information, see [Addressing cell's context parameter](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter).</span></span> <span data-ttu-id="96f3d-182">不能将自定义函数同时设置为流式处理和 requiresAddress。</span><span class="sxs-lookup"><span data-stu-id="96f3d-182">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="96f3d-183">使用此选项时，"调用" 参数必须是在 options 中传递的最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="96f3d-183">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span>                                              |
| `stream`          | <span data-ttu-id="96f3d-184">boolean</span><span class="sxs-lookup"><span data-stu-id="96f3d-184">boolean</span></span>   | <span data-ttu-id="96f3d-185">否</span><span class="sxs-lookup"><span data-stu-id="96f3d-185">No</span></span><br/><br/><span data-ttu-id="96f3d-186">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="96f3d-186">Default value is `false`.</span></span>  | <span data-ttu-id="96f3d-187">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="96f3d-187">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="96f3d-188">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="96f3d-188">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="96f3d-189">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="96f3d-189">The function should have no `return` statement.</span></span> <span data-ttu-id="96f3d-190">相反，结果值将作为 `StreamingInvocation.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="96f3d-190">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="96f3d-191">有关详细信息，请参阅[流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)。</span><span class="sxs-lookup"><span data-stu-id="96f3d-191">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>                                                                                                                                                                |
| `volatile`        | <span data-ttu-id="96f3d-192">boolean</span><span class="sxs-lookup"><span data-stu-id="96f3d-192">boolean</span></span>   | <span data-ttu-id="96f3d-193">否</span><span class="sxs-lookup"><span data-stu-id="96f3d-193">No</span></span> <br/><br/><span data-ttu-id="96f3d-194">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="96f3d-194">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="96f3d-195">如果为 `true`，则该函数会在每次 Excel 重新计算时（而不是仅当公式的从属值发生更改时）进行重新计算。</span><span class="sxs-lookup"><span data-stu-id="96f3d-195">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="96f3d-196">函数不能同时为流式处理和可变。</span><span class="sxs-lookup"><span data-stu-id="96f3d-196">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="96f3d-197">如果 `stream` 和 `volatile` 属性同时设置为 `true`，则将忽略可变选项。</span><span class="sxs-lookup"><span data-stu-id="96f3d-197">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span>                                                                                                                                                                                                                                                                                             |

### <a name="parameters"></a><span data-ttu-id="96f3d-198">参数</span><span class="sxs-lookup"><span data-stu-id="96f3d-198">parameters</span></span>

<span data-ttu-id="96f3d-199">`parameters` 属性是参数对象的数组。</span><span class="sxs-lookup"><span data-stu-id="96f3d-199">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="96f3d-200">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="96f3d-200">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="96f3d-201">属性</span><span class="sxs-lookup"><span data-stu-id="96f3d-201">Property</span></span>  |  <span data-ttu-id="96f3d-202">数据类型</span><span class="sxs-lookup"><span data-stu-id="96f3d-202">Data type</span></span>  |  <span data-ttu-id="96f3d-203">必需</span><span class="sxs-lookup"><span data-stu-id="96f3d-203">Required</span></span>  |  <span data-ttu-id="96f3d-204">说明</span><span class="sxs-lookup"><span data-stu-id="96f3d-204">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="96f3d-205">string</span><span class="sxs-lookup"><span data-stu-id="96f3d-205">string</span></span>  |  <span data-ttu-id="96f3d-206">No</span><span class="sxs-lookup"><span data-stu-id="96f3d-206">No</span></span> |  <span data-ttu-id="96f3d-207">参数的说明。</span><span class="sxs-lookup"><span data-stu-id="96f3d-207">A description of the parameter.</span></span> <span data-ttu-id="96f3d-208">这显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="96f3d-208">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="96f3d-209">string</span><span class="sxs-lookup"><span data-stu-id="96f3d-209">string</span></span>  |  <span data-ttu-id="96f3d-210">否</span><span class="sxs-lookup"><span data-stu-id="96f3d-210">No</span></span>  |  <span data-ttu-id="96f3d-211">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="96f3d-211">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="96f3d-212">string</span><span class="sxs-lookup"><span data-stu-id="96f3d-212">string</span></span>  |  <span data-ttu-id="96f3d-213">是</span><span class="sxs-lookup"><span data-stu-id="96f3d-213">Yes</span></span>  |  <span data-ttu-id="96f3d-214">参数的名称。</span><span class="sxs-lookup"><span data-stu-id="96f3d-214">The name of the parameter.</span></span> <span data-ttu-id="96f3d-215">此名称显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="96f3d-215">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="96f3d-216">string</span><span class="sxs-lookup"><span data-stu-id="96f3d-216">string</span></span>  |  <span data-ttu-id="96f3d-217">No</span><span class="sxs-lookup"><span data-stu-id="96f3d-217">No</span></span>  |  <span data-ttu-id="96f3d-218">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="96f3d-218">The data type of the parameter.</span></span> <span data-ttu-id="96f3d-219">可以是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。</span><span class="sxs-lookup"><span data-stu-id="96f3d-219">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="96f3d-220">如果未指定此属性，则数据类型默认为 **any**。</span><span class="sxs-lookup"><span data-stu-id="96f3d-220">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="96f3d-221">boolean</span><span class="sxs-lookup"><span data-stu-id="96f3d-221">boolean</span></span> | <span data-ttu-id="96f3d-222">否</span><span class="sxs-lookup"><span data-stu-id="96f3d-222">No</span></span> | <span data-ttu-id="96f3d-223">如果为 `true`，则参数是可选的。</span><span class="sxs-lookup"><span data-stu-id="96f3d-223">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="96f3d-224">boolean</span><span class="sxs-lookup"><span data-stu-id="96f3d-224">boolean</span></span> | <span data-ttu-id="96f3d-225">否</span><span class="sxs-lookup"><span data-stu-id="96f3d-225">No</span></span> | <span data-ttu-id="96f3d-226">如果`true`参数，则从指定的数组填充。</span><span class="sxs-lookup"><span data-stu-id="96f3d-226">If `true`, parameters will populate from a specified array.</span></span> <span data-ttu-id="96f3d-227">请注意，根据定义，所有重复参数均被视为可选参数。</span><span class="sxs-lookup"><span data-stu-id="96f3d-227">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="96f3d-228">结果</span><span class="sxs-lookup"><span data-stu-id="96f3d-228">result</span></span>

<span data-ttu-id="96f3d-229">`result` 对象定义函数返回的信息类型。</span><span class="sxs-lookup"><span data-stu-id="96f3d-229">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="96f3d-230">下表列出了 `result` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="96f3d-230">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="96f3d-231">属性</span><span class="sxs-lookup"><span data-stu-id="96f3d-231">Property</span></span>         | <span data-ttu-id="96f3d-232">数据类型</span><span class="sxs-lookup"><span data-stu-id="96f3d-232">Data type</span></span> | <span data-ttu-id="96f3d-233">必需</span><span class="sxs-lookup"><span data-stu-id="96f3d-233">Required</span></span> | <span data-ttu-id="96f3d-234">说明</span><span class="sxs-lookup"><span data-stu-id="96f3d-234">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="96f3d-235">string</span><span class="sxs-lookup"><span data-stu-id="96f3d-235">string</span></span>    | <span data-ttu-id="96f3d-236">No</span><span class="sxs-lookup"><span data-stu-id="96f3d-236">No</span></span>       | <span data-ttu-id="96f3d-237">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="96f3d-237">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="96f3d-238">将函数名称与 JSON 元数据相关联</span><span class="sxs-lookup"><span data-stu-id="96f3d-238">Associating function names with JSON metadata</span></span>

<span data-ttu-id="96f3d-239">若要使函数正常工作，需要将函数的`id`属性与 JavaScript 实现相关联。</span><span class="sxs-lookup"><span data-stu-id="96f3d-239">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="96f3d-240">请确保存在关联，否则将不会在 Excel 中注册该函数，也不能使用它。</span><span class="sxs-lookup"><span data-stu-id="96f3d-240">Make sure there is an association, otherwise the function will not be registered and not useable in Excel.</span></span> <span data-ttu-id="96f3d-241">下面的代码示例演示如何使用`CustomFunctions.associate()`方法进行关联。</span><span class="sxs-lookup"><span data-stu-id="96f3d-241">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="96f3d-242">该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。</span><span class="sxs-lookup"><span data-stu-id="96f3d-242">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="96f3d-243">下面的 JSON 显示了与上一个自定义函数 JavaScript 代码相关联的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="96f3d-243">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

```json
{
  "functions": [
    {
      "description": "Add two numbers",
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        {
          "description": "First number",
          "name": "first",
          "type": "number"
        },
        {
          "description": "Second number",
          "name": "second",
          "type": "number"
        }
      ],
      "result": {
        "type": "number"
      }
    }
  ]
}
```

<span data-ttu-id="96f3d-244">在 JavaScript 文件中创建自定义函数和在 JSON 元数据文件中指定相应信息时，请记住以下最佳实践。</span><span class="sxs-lookup"><span data-stu-id="96f3d-244">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="96f3d-245">在 JSON 元数据文件中，确保每个 `id` 属性的值仅包含字母数字字符和句点。</span><span class="sxs-lookup"><span data-stu-id="96f3d-245">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="96f3d-246">在 JSON 元数据文件中，确保每个 `id` 属性的值在该文件范围内是唯一的。</span><span class="sxs-lookup"><span data-stu-id="96f3d-246">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="96f3d-247">也就是说，元数据文件中不应存在具有相同 `id` 值的两个函数对象。</span><span class="sxs-lookup"><span data-stu-id="96f3d-247">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="96f3d-248">在将 JSON 元数据文件中的 `id` 属性的值与相应的 JavaScript 函数名称关联后，请勿再更改该值。</span><span class="sxs-lookup"><span data-stu-id="96f3d-248">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="96f3d-249">你可以通过更新 JSON 元数据文件中的 `name` 属性来更改最终用户在 Excel 中看到的函数名称，但绝不能更改已确定的 `id` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="96f3d-249">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="96f3d-250">在 JavaScript 文件中，使用`CustomFunctions.associate`每个函数的后面指定自定义函数关联。</span><span class="sxs-lookup"><span data-stu-id="96f3d-250">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="96f3d-251">以下示例显示了与此 JavaScript 代码示例中定义的函数相对应的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="96f3d-251">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="96f3d-252">`id`和`name`属性值以大写形式表示，这是描述自定义函数的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="96f3d-252">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="96f3d-253">仅当您手动准备自己的 JSON 文件，而不是使用自动生成时，才需要添加此 JSON。</span><span class="sxs-lookup"><span data-stu-id="96f3d-253">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="96f3d-254">有关自动生成的详细信息，请参阅[CREATE JSON metadata for custom 函数](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="96f3d-254">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="next-steps"></a><span data-ttu-id="96f3d-255">后续步骤</span><span class="sxs-lookup"><span data-stu-id="96f3d-255">Next steps</span></span>

<span data-ttu-id="96f3d-256">了解[有关命名函数](custom-functions-naming.md)或了解如何使用前面所述的手写 JSON 方法对[函数进行本地化](custom-functions-localize.md)的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="96f3d-256">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="96f3d-257">另请参阅</span><span class="sxs-lookup"><span data-stu-id="96f3d-257">See also</span></span>

- [<span data-ttu-id="96f3d-258">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="96f3d-258">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="96f3d-259">自定义函数参数选项</span><span class="sxs-lookup"><span data-stu-id="96f3d-259">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="96f3d-260">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="96f3d-260">Create custom functions in Excel</span></span>](custom-functions-overview.md)

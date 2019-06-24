---
ms.date: 06/18/2019
description: 了解在 Excel 中开发自定义函数的最佳实践。
title: 自定义函数最佳实践
localization_priority: Normal
ms.openlocfilehash: 7c836119a783f5cc7e1e7f4f52f1d21b86091bfe
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127931"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="73dee-103">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="73dee-103">Custom functions best practices</span></span>

<span data-ttu-id="73dee-104">本文介绍了在 Excel 中开发自定义函数的最佳实践。</span><span class="sxs-lookup"><span data-stu-id="73dee-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="73dee-105">将函数名称与 JSON 元数据相关联</span><span class="sxs-lookup"><span data-stu-id="73dee-105">Associating function names with JSON metadata</span></span>

<span data-ttu-id="73dee-106">如[自定义函数概述](custom-functions-overview.md)文章中所述，自定义函数项目必须包含 JSON 元数据文件和脚本（JavaScript 或 TypeScript）文件才能构成完整的函数。</span><span class="sxs-lookup"><span data-stu-id="73dee-106">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="73dee-107">如果您使用`yo office`的是 JSON 元数据, 则可以从代码注释生成。</span><span class="sxs-lookup"><span data-stu-id="73dee-107">If you are using `yo office` the JSON metadata can be generated from the code comments.</span></span> <span data-ttu-id="73dee-108">否则, 您需要手动生成 JSON 元数据文件。</span><span class="sxs-lookup"><span data-stu-id="73dee-108">Otherwise you need to build the JSON metadata file manually.</span></span>

<span data-ttu-id="73dee-109">若要使函数正常工作, 需要将函数的`id`属性与 JavaScript 实现相关联。</span><span class="sxs-lookup"><span data-stu-id="73dee-109">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="73dee-110">请确保存在关联, 否则将不会调用该函数。</span><span class="sxs-lookup"><span data-stu-id="73dee-110">Make sure there is an association, otherwise the function will not be called.</span></span> <span data-ttu-id="73dee-111">下面的代码示例演示如何使用`CustomFunctions.associate()`方法进行关联。</span><span class="sxs-lookup"><span data-stu-id="73dee-111">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="73dee-112">该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。</span><span class="sxs-lookup"><span data-stu-id="73dee-112">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="73dee-113">下面的 JSON 显示了与上一个自定义函数 JavaScript 代码相关联的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="73dee-113">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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
    },
  ]
}
```


<span data-ttu-id="73dee-114">在 JavaScript 文件中创建自定义函数和在 JSON 元数据文件中指定相应信息时，请记住以下最佳实践。</span><span class="sxs-lookup"><span data-stu-id="73dee-114">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="73dee-115">在 JSON 元数据文件中，确保每个 `id` 属性的值仅包含字母数字字符和句点。</span><span class="sxs-lookup"><span data-stu-id="73dee-115">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="73dee-116">在 JSON 元数据文件中，确保每个 `id` 属性的值在该文件范围内是唯一的。</span><span class="sxs-lookup"><span data-stu-id="73dee-116">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="73dee-117">也就是说，元数据文件中不应存在具有相同 `id` 值的两个函数对象。</span><span class="sxs-lookup"><span data-stu-id="73dee-117">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

* <span data-ttu-id="73dee-118">在将 JSON 元数据文件中的 `id` 属性的值与相应的 JavaScript 函数名称关联后，请勿再更改该值。</span><span class="sxs-lookup"><span data-stu-id="73dee-118">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="73dee-119">你可以通过更新 JSON 元数据文件中的 `name` 属性来更改最终用户在 Excel 中看到的函数名称，但绝不能更改已确定的 `id` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="73dee-119">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="73dee-120">在 JavaScript 文件中, 使用`CustomFunctions.associate`每个函数的后面指定自定义函数关联。</span><span class="sxs-lookup"><span data-stu-id="73dee-120">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="73dee-121">以下示例显示了与此 JavaScript 代码示例中定义的函数相对应的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="73dee-121">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="73dee-122">`id`和`name`属性值以大写形式表示, 这是描述自定义函数的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="73dee-122">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="73dee-123">仅当您手动准备自己的 JSON 文件, 而不是使用自动生成时, 才需要添加此 JSON。</span><span class="sxs-lookup"><span data-stu-id="73dee-123">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="73dee-124">有关自动生成的详细信息, 请参阅[CREATE JSON metadata for custom 函数](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="73dee-124">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="73dee-125">其他注意事项</span><span class="sxs-lookup"><span data-stu-id="73dee-125">Additional considerations</span></span>

<span data-ttu-id="73dee-126">避免从自定义函数中直接或间接访问文档对象模型 (DOM) (例如, 使用 jQuery)。</span><span class="sxs-lookup"><span data-stu-id="73dee-126">Avoid accessing the Document Object Model (DOM) directly or indirectly (for example, using jQuery) from your custom function.</span></span> <span data-ttu-id="73dee-127">在 Windows 的 Excel 中, 自定义函数使用[JavaScript 运行时](custom-functions-runtime.md), 自定义函数无法访问 DOM。</span><span class="sxs-lookup"><span data-stu-id="73dee-127">In Excel on Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="73dee-128">后续步骤</span><span class="sxs-lookup"><span data-stu-id="73dee-128">Next steps</span></span>
<span data-ttu-id="73dee-129">了解如何[使用自定义函数执行 web 请求](custom-functions-web-reqs.md)。</span><span class="sxs-lookup"><span data-stu-id="73dee-129">Learn how to [perform web requests with custom functions](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="73dee-130">另请参阅</span><span class="sxs-lookup"><span data-stu-id="73dee-130">See also</span></span>

* [<span data-ttu-id="73dee-131">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="73dee-131">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="73dee-132">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="73dee-132">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="73dee-133">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="73dee-133">Create custom functions in Excel</span></span>](custom-functions-overview.md)

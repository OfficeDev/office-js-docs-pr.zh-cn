---
ms.date: 03/15/2021
description: 使用 JSDoc 标记动态创建自定义函数 JSON 元数据。
title: 为自定义函数自动生成 JSON 元数据
localization_priority: Normal
ms.openlocfilehash: e31059de78e9daedc31c9b0a8605b5352fd0ed94
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178046"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="d6783-103">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="d6783-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="d6783-104">在使用 JavaScript 或 TypeScript 编写 Excel 自定义函数时，使用 [JSDoc 标记](https://jsdoc.app/)提供有关自定义函数的额外信息。</span><span class="sxs-lookup"><span data-stu-id="d6783-104">When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function.</span></span> <span data-ttu-id="d6783-105">然后在生成时使用 JSDoc 标记创建 JSON 元数据文件。</span><span class="sxs-lookup"><span data-stu-id="d6783-105">The JSDoc tags are then used at build time to create the JSON metadata file.</span></span> <span data-ttu-id="d6783-106">使用 JSDoc 标记，你无需手动编辑 [JSON 元数据文件](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="d6783-106">Using JSDoc tags saves you from the effort of [manually editing the JSON metadata file](custom-functions-json.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="d6783-107">为 JavaScript 或 TypeScript 函数添加代码注释中的 `@customfunction` 标记以将其标记为自定义函数。</span><span class="sxs-lookup"><span data-stu-id="d6783-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="d6783-108">可以使用 JavaScript 中的 [@param](#param) 标记或从 TypeScript 中的[函数类型](https://www.typescriptlang.org/docs/handbook/functions.html)提供函数参数类型。</span><span class="sxs-lookup"><span data-stu-id="d6783-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="d6783-109">有关详细信息，请参阅 [@param](#param) 标记和[类型](#types)部分。</span><span class="sxs-lookup"><span data-stu-id="d6783-109">For more information, see the [@param](#param) tag and [Types](#types) sections.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="d6783-110">为函数添加说明</span><span class="sxs-lookup"><span data-stu-id="d6783-110">Adding a description to a function</span></span>

<span data-ttu-id="d6783-111">当用户需要帮助来了解自定义函数的功能时，将向用户显示用作帮助文本的说明。</span><span class="sxs-lookup"><span data-stu-id="d6783-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="d6783-112">说明不需要任何特定标记。</span><span class="sxs-lookup"><span data-stu-id="d6783-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="d6783-113">只需在 JSDoc 注释中输入简短的文本说明即可。</span><span class="sxs-lookup"><span data-stu-id="d6783-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="d6783-114">一般来说，说明位于 JSDoc 注释部分的开头，但无论位于何处，它都有用。</span><span class="sxs-lookup"><span data-stu-id="d6783-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="d6783-115">若要查看内置函数说明的示例，请打开 Excel，转到“**公式**”选项卡，然后选择“**插入函数**”。</span><span class="sxs-lookup"><span data-stu-id="d6783-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="d6783-116">然后，你可以浏览所有函数说明，还可以查看列出的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="d6783-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="d6783-117">在以下示例中，短语“计算球体的体积”</span><span class="sxs-lookup"><span data-stu-id="d6783-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="d6783-118">就是自定义函数的相关说明。</span><span class="sxs-lookup"><span data-stu-id="d6783-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="d6783-119">JSDoc 标记</span><span class="sxs-lookup"><span data-stu-id="d6783-119">JSDoc Tags</span></span>

<span data-ttu-id="d6783-120">以下 JSDoc 标记在 Excel 自定义函数中受支持。</span><span class="sxs-lookup"><span data-stu-id="d6783-120">The following JSDoc tags are supported in Excel custom functions.</span></span>

* [<span data-ttu-id="d6783-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="d6783-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="d6783-122">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="d6783-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="d6783-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="d6783-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="d6783-124">[@param](#param) _{type}_ name description</span><span class="sxs-lookup"><span data-stu-id="d6783-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="d6783-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="d6783-125">@requiresAddress</span></span>](#requiresAddress)
* [<span data-ttu-id="d6783-126">@requiresParameterAddresses</span><span class="sxs-lookup"><span data-stu-id="d6783-126">@requiresParameterAddresses</span></span>](#requiresParameterAddresses)
* <span data-ttu-id="d6783-127">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="d6783-127">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="d6783-128">@streaming</span><span class="sxs-lookup"><span data-stu-id="d6783-128">@streaming</span></span>](#streaming)
* [<span data-ttu-id="d6783-129">@volatile</span><span class="sxs-lookup"><span data-stu-id="d6783-129">@volatile</span></span>](#volatile)

---
<a id="cancelable"></a>
### <a name="cancelable"></a><span data-ttu-id="d6783-130">@cancelable</span><span class="sxs-lookup"><span data-stu-id="d6783-130">@cancelable</span></span>

<span data-ttu-id="d6783-131">指示自定义函数在函数取消时执行一个操作。</span><span class="sxs-lookup"><span data-stu-id="d6783-131">Indicates that a custom function performs an action when the function is canceled.</span></span>

<span data-ttu-id="d6783-132">最后一个函数参数的类型必须是 `CustomFunctions.CancelableInvocation`。</span><span class="sxs-lookup"><span data-stu-id="d6783-132">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="d6783-133">函数可以将函数分配给 `oncanceled` 属性，以在函数取消时表示结果。</span><span class="sxs-lookup"><span data-stu-id="d6783-133">The function can assign a function to the `oncanceled` property to denote the result when the function is canceled.</span></span>

<span data-ttu-id="d6783-134">如果最后一个函数参数的类型为 `CustomFunctions.CancelableInvocation`，则即使标记不存在，也会被视为 `@cancelable`。</span><span class="sxs-lookup"><span data-stu-id="d6783-134">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.</span></span>

<span data-ttu-id="d6783-135">函数不能同时具有 `@cancelable` 和 `@streaming` 标记。</span><span class="sxs-lookup"><span data-stu-id="d6783-135">A function can't have both `@cancelable` and `@streaming` tags.</span></span>

<a id="customfunction"></a>

### <a name="customfunction"></a><span data-ttu-id="d6783-136">@customfunction</span><span class="sxs-lookup"><span data-stu-id="d6783-136">@customfunction</span></span>

<span data-ttu-id="d6783-137">语法：@customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="d6783-137">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="d6783-138">此标记指示 JavaScript/TypeScript 函数是 Excel 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="d6783-138">This tag indicates that the JavaScript/TypeScript function is an Excel custom function.</span></span> <span data-ttu-id="d6783-139">需要为自定义函数创建元数据。</span><span class="sxs-lookup"><span data-stu-id="d6783-139">It is required to create metadata for the custom function.</span></span>

<span data-ttu-id="d6783-140">下面显示了此标记的示例。</span><span class="sxs-lookup"><span data-stu-id="d6783-140">The following shows an example of this tag.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="d6783-141">id</span><span class="sxs-lookup"><span data-stu-id="d6783-141">id</span></span>

<span data-ttu-id="d6783-142">`id`标识自定义函数。</span><span class="sxs-lookup"><span data-stu-id="d6783-142">The `id` identifies a custom function.</span></span>

* <span data-ttu-id="d6783-143">如果未提供 `id`，请将 JavaScript/TypeScript 函数名称转换为大写并删除禁用字符。</span><span class="sxs-lookup"><span data-stu-id="d6783-143">If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.</span></span>
* <span data-ttu-id="d6783-144">`id` 对于所有自定义函数必须是唯一的。</span><span class="sxs-lookup"><span data-stu-id="d6783-144">The `id` must be unique for all custom functions.</span></span>
* <span data-ttu-id="d6783-145">允许使用的字符限为：A-Z、a-z、0-9、下划线 (\_) 和句点 (.)。</span><span class="sxs-lookup"><span data-stu-id="d6783-145">The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).</span></span>

<span data-ttu-id="d6783-146">在下面的示例中，增量是函数的 `id` 和 `name`。</span><span class="sxs-lookup"><span data-stu-id="d6783-146">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="d6783-147">name</span><span class="sxs-lookup"><span data-stu-id="d6783-147">name</span></span>

<span data-ttu-id="d6783-148">提供自定义函数的显示`name`。</span><span class="sxs-lookup"><span data-stu-id="d6783-148">Provides the display `name` for the custom function.</span></span>

* <span data-ttu-id="d6783-149">如果未提供名称，则 id 还会用作名称。</span><span class="sxs-lookup"><span data-stu-id="d6783-149">If name isn't provided, the id is also used as the name.</span></span>
* <span data-ttu-id="d6783-150">允许使用的字符：字母 [Unicode 字母字符](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、句点 (.) 和下划线 (\_)。</span><span class="sxs-lookup"><span data-stu-id="d6783-150">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="d6783-151">必须以字母开头。</span><span class="sxs-lookup"><span data-stu-id="d6783-151">Must start with a letter.</span></span>
* <span data-ttu-id="d6783-152">最大长度为 128 个字符。</span><span class="sxs-lookup"><span data-stu-id="d6783-152">Maximum length is 128 characters.</span></span>

<span data-ttu-id="d6783-153">在下面的示例中，INC 是函数的 `id`，并且 `increment` 是 `name`。</span><span class="sxs-lookup"><span data-stu-id="d6783-153">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="d6783-154">说明</span><span class="sxs-lookup"><span data-stu-id="d6783-154">description</span></span>

<span data-ttu-id="d6783-155">Excel 中的用户在输入函数时会显示说明，并指定函数的功能。</span><span class="sxs-lookup"><span data-stu-id="d6783-155">A description appears to users in Excel as they are entering the function and specifies what the function does.</span></span> <span data-ttu-id="d6783-156">说明不需要任何特定标记。</span><span class="sxs-lookup"><span data-stu-id="d6783-156">A description doesn't require any specific tag.</span></span> <span data-ttu-id="d6783-157">通过在 JSDoc 注释中添加一个短语来描述函数的功能，为自定义函数添加说明。</span><span class="sxs-lookup"><span data-stu-id="d6783-157">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="d6783-158">默认情况下，JSDoc 注释部分中未标记的任何文本都是该函数的说明。</span><span class="sxs-lookup"><span data-stu-id="d6783-158">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span>

<span data-ttu-id="d6783-159">在以下示例中，短语“对两个数字求和的函数”是 id 属性为 `ADD` 的自定义函数的相关说明。</span><span class="sxs-lookup"><span data-stu-id="d6783-159">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="helpurl"></a>
### <a name="helpurl"></a><span data-ttu-id="d6783-160">@helpurl</span><span class="sxs-lookup"><span data-stu-id="d6783-160">@helpurl</span></span>

<span data-ttu-id="d6783-161">语法：@helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="d6783-161">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="d6783-162">提供的 _url_ 显示在 Excel 中。</span><span class="sxs-lookup"><span data-stu-id="d6783-162">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="d6783-163">在下面的示例中，为 `helpurl` `www.contoso.com/weatherhelp` 。</span><span class="sxs-lookup"><span data-stu-id="d6783-163">In the following example, the `helpurl` is `www.contoso.com/weatherhelp`.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

<a id="param"></a>
### <a name="param"></a><span data-ttu-id="d6783-164">@param</span><span class="sxs-lookup"><span data-stu-id="d6783-164">@param</span></span>

#### <a name="javascript"></a><span data-ttu-id="d6783-165">JavaScript</span><span class="sxs-lookup"><span data-stu-id="d6783-165">JavaScript</span></span>

<span data-ttu-id="d6783-166">JavaScript 语法：@param {type} name _description_</span><span class="sxs-lookup"><span data-stu-id="d6783-166">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="d6783-167">`{type}` 指定大括号中的类型信息。</span><span class="sxs-lookup"><span data-stu-id="d6783-167">`{type}` specifies the type info within curly braces.</span></span> <span data-ttu-id="d6783-168">有关可能使用的类型的详细信息，请参阅[类型](#types)部分。</span><span class="sxs-lookup"><span data-stu-id="d6783-168">See the [Types](#types) section for more information about the types which may be used.</span></span> <span data-ttu-id="d6783-169">如果未指定类型，则使用 `any` 默认类型。</span><span class="sxs-lookup"><span data-stu-id="d6783-169">If no type is specified, the default type `any` will be used.</span></span>
* <span data-ttu-id="d6783-170">`name` 指定该标记@param参数。</span><span class="sxs-lookup"><span data-stu-id="d6783-170">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="d6783-171">这是必需的。</span><span class="sxs-lookup"><span data-stu-id="d6783-171">It is required.</span></span>
* <span data-ttu-id="d6783-172">`description` 为函数参数提供显示在 Excel 中的说明。</span><span class="sxs-lookup"><span data-stu-id="d6783-172">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="d6783-173">可选。</span><span class="sxs-lookup"><span data-stu-id="d6783-173">It is optional.</span></span>

<span data-ttu-id="d6783-174">若要将自定义函数参数表示为可选，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="d6783-174">To denote a custom function parameter as optional:</span></span>

* <span data-ttu-id="d6783-175">为参数名称加上方括号。</span><span class="sxs-lookup"><span data-stu-id="d6783-175">Put square brackets around the parameter name.</span></span> <span data-ttu-id="d6783-176">例如：`@param {string} [text] Optional text`。</span><span class="sxs-lookup"><span data-stu-id="d6783-176">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="d6783-177">可选参数的默认值为 `null`。</span><span class="sxs-lookup"><span data-stu-id="d6783-177">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="d6783-178">以下示例显示添加两个或三个数字的 ADD 函数，第三个数字作为可选参数。</span><span class="sxs-lookup"><span data-stu-id="d6783-178">The following example shows an ADD function that adds two or three numbers, with the third number as an optional parameter.</span></span>

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a><span data-ttu-id="d6783-179">TypeScript</span><span class="sxs-lookup"><span data-stu-id="d6783-179">TypeScript</span></span>

<span data-ttu-id="d6783-180">TypeScript 语法：@param name _description_</span><span class="sxs-lookup"><span data-stu-id="d6783-180">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="d6783-181">`name` 指定该标记@param参数。</span><span class="sxs-lookup"><span data-stu-id="d6783-181">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="d6783-182">这是必需的。</span><span class="sxs-lookup"><span data-stu-id="d6783-182">It is required.</span></span>
* <span data-ttu-id="d6783-183">`description` 为函数参数提供显示在 Excel 中的说明。</span><span class="sxs-lookup"><span data-stu-id="d6783-183">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="d6783-184">可选。</span><span class="sxs-lookup"><span data-stu-id="d6783-184">It is optional.</span></span>

<span data-ttu-id="d6783-185">有关可能使用的函数参数类型的详细信息，请参阅[类型](#types)部分。</span><span class="sxs-lookup"><span data-stu-id="d6783-185">See the [Types](#types) section for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="d6783-186">若要将自定义函数参数表示为可选，请执行以下操作之一：</span><span class="sxs-lookup"><span data-stu-id="d6783-186">To denote a custom function parameter as optional, do one of the following:</span></span>

* <span data-ttu-id="d6783-187">使用可选参数。</span><span class="sxs-lookup"><span data-stu-id="d6783-187">Use an optional parameter.</span></span> <span data-ttu-id="d6783-188">例如：`function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="d6783-188">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="d6783-189">为该参数提供默认值。</span><span class="sxs-lookup"><span data-stu-id="d6783-189">Give the parameter a default value.</span></span> <span data-ttu-id="d6783-190">例如：`function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="d6783-190">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="d6783-191">有关 @param 的详细说明，请参阅：[JSDoc](https://jsdoc.app/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="d6783-191">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="d6783-192">可选参数的默认值为 `null`。</span><span class="sxs-lookup"><span data-stu-id="d6783-192">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="d6783-193">下面的示例显示了将两个数字相加的 `add` 函数。</span><span class="sxs-lookup"><span data-stu-id="d6783-193">The following example shows the `add` function that adds two numbers.</span></span>

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="requiresAddress"></a>

### <a name="requiresaddress"></a><span data-ttu-id="d6783-194">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="d6783-194">@requiresAddress</span></span>

<span data-ttu-id="d6783-195">表示应提供计算函数所在的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="d6783-195">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="d6783-196">最后一个函数参数必须为 类型 `CustomFunctions.Invocation` 或派生类型，以使用 `@requiresAddress` 。</span><span class="sxs-lookup"><span data-stu-id="d6783-196">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type to use `@requiresAddress`.</span></span> <span data-ttu-id="d6783-197">调用函数时，`address` 属性将包含地址。</span><span class="sxs-lookup"><span data-stu-id="d6783-197">When the function is called, the `address` property will contain the address.</span></span>

<span data-ttu-id="d6783-198">以下示例演示如何将 参数与 结合使用以返回调用自定义函数 `invocation` `@requiresAddress` 的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="d6783-198">The following sample shows how to use the `invocation` parameter in combination with `@requiresAddress` to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="d6783-199">有关详细信息 [，请参阅调用](custom-functions-parameter-options.md#invocation-parameter) 参数。</span><span class="sxs-lookup"><span data-stu-id="d6783-199">See [Invocation parameter](custom-functions-parameter-options.md#invocation-parameter) for more information.</span></span>

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<a id="requiresParameterAddresses"></a>
### <a name="requiresparameteraddresses"></a><span data-ttu-id="d6783-200">@requiresParameterAddresses</span><span class="sxs-lookup"><span data-stu-id="d6783-200">@requiresParameterAddresses</span></span>

<span data-ttu-id="d6783-201">指示函数应返回输入参数的地址。</span><span class="sxs-lookup"><span data-stu-id="d6783-201">Indicates that the function should return the addresses of input parameters.</span></span> 

<span data-ttu-id="d6783-202">最后一个函数参数必须为 类型 `CustomFunctions.Invocation` 或派生类型，以使用  `@requiresParameterAddresses` 。</span><span class="sxs-lookup"><span data-stu-id="d6783-202">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type to use  `@requiresParameterAddresses`.</span></span> <span data-ttu-id="d6783-203">JSDoc 注释还必须包含一个标记，该标记指定返回 `@returns` 值是矩阵，如 `@returns {string[][]}` 或 `@returns {number[][]}` 。</span><span class="sxs-lookup"><span data-stu-id="d6783-203">The JSDoc comment must also include an `@returns` tag specifying that the return value be a matrix, such as `@returns {string[][]}` or `@returns {number[][]}`.</span></span> <span data-ttu-id="d6783-204">有关 [其他信息，](#matrix-type) 请参阅矩阵类型。</span><span class="sxs-lookup"><span data-stu-id="d6783-204">See [Matrix types](#matrix-type) for additional information.</span></span> 

<span data-ttu-id="d6783-205">调用 函数时， `parameterAddresses` 属性将包含输入参数的地址。</span><span class="sxs-lookup"><span data-stu-id="d6783-205">When the function is called, the `parameterAddresses` property will contain the addresses of the input parameters.</span></span>

<span data-ttu-id="d6783-206">以下示例演示如何将 参数与 结合使用以返回三个 `invocation` `@requiresParameterAddresses` 输入参数的地址。</span><span class="sxs-lookup"><span data-stu-id="d6783-206">The following sample shows how to use the `invocation` parameter in combination with `@requiresParameterAddresses` to return the addresses of three input parameters.</span></span> <span data-ttu-id="d6783-207">有关详细信息 [，请参阅检测参数](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) 的地址。</span><span class="sxs-lookup"><span data-stu-id="d6783-207">See [Detect the address of a parameter](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) for more information.</span></span> 

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array.
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<a id="returns"></a>
### <a name="returns"></a><span data-ttu-id="d6783-208">@returns</span><span class="sxs-lookup"><span data-stu-id="d6783-208">@returns</span></span>

<span data-ttu-id="d6783-209">语法：@returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="d6783-209">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="d6783-210">提供返回值的类型。</span><span class="sxs-lookup"><span data-stu-id="d6783-210">Provides the type for the return value.</span></span>

<span data-ttu-id="d6783-211">如果省略 `{type}`，则将使用 TypeScript 类型信息。</span><span class="sxs-lookup"><span data-stu-id="d6783-211">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="d6783-212">如果没有类型信息，则类型将为 `any`。</span><span class="sxs-lookup"><span data-stu-id="d6783-212">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="d6783-213">下面的示例显示了使用 `@returns` 标记的 `add` 函数。</span><span class="sxs-lookup"><span data-stu-id="d6783-213">The following example shows the `add` function that uses the `@returns` tag.</span></span>

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="streaming"></a>
### <a name="streaming"></a><span data-ttu-id="d6783-214">@streaming</span><span class="sxs-lookup"><span data-stu-id="d6783-214">@streaming</span></span>

<span data-ttu-id="d6783-215">用于表示自定义函数是一个流式处理函数。</span><span class="sxs-lookup"><span data-stu-id="d6783-215">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="d6783-216">最后一个参数的类型为 `CustomFunctions.StreamingInvocation<ResultType>` 。</span><span class="sxs-lookup"><span data-stu-id="d6783-216">The last parameter is of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="d6783-217">函数返回 `void` 。</span><span class="sxs-lookup"><span data-stu-id="d6783-217">The function returns `void`.</span></span>

<span data-ttu-id="d6783-218">流式处理函数不会直接返回值，而是使用 `setResult(result: ResultType)` 最后一个参数调用。</span><span class="sxs-lookup"><span data-stu-id="d6783-218">Streaming functions don't return values directly, instead they call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="d6783-219">由流式处理函数引发的异常将被忽略。</span><span class="sxs-lookup"><span data-stu-id="d6783-219">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="d6783-220">`setResult()` 可能称为“错误”，以指示错误结果。</span><span class="sxs-lookup"><span data-stu-id="d6783-220">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="d6783-221">有关流式处理函数的示例和更多信息，请参阅[生成流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)。</span><span class="sxs-lookup"><span data-stu-id="d6783-221">For an example of a streaming function and more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="d6783-222">流式处理函数不能标记为 [@volatile](#volatile)。</span><span class="sxs-lookup"><span data-stu-id="d6783-222">Streaming functions can't be marked as [@volatile](#volatile).</span></span>

<a id="volatile"></a>
### <a name="volatile"></a><span data-ttu-id="d6783-223">@volatile</span><span class="sxs-lookup"><span data-stu-id="d6783-223">@volatile</span></span>

<span data-ttu-id="d6783-224">可变函数是指其结果不断变化的函数，即使不采用任何参数或参数未发生更改都是如此。</span><span class="sxs-lookup"><span data-stu-id="d6783-224">A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed.</span></span> <span data-ttu-id="d6783-225">Excel 在每次完成计算后，都会重新计算包含可变函数和所有依赖项的单元格。</span><span class="sxs-lookup"><span data-stu-id="d6783-225">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="d6783-226">因此，过于依赖可变函数会使重新计算时间变慢，请谨慎使用。</span><span class="sxs-lookup"><span data-stu-id="d6783-226">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="d6783-227">流式处理函数不能为可变函数。</span><span class="sxs-lookup"><span data-stu-id="d6783-227">Streaming functions can't be volatile.</span></span>

<span data-ttu-id="d6783-228">以下函数是可变函数并使用 `@volatile` 标记。</span><span class="sxs-lookup"><span data-stu-id="d6783-228">The following function is volatile and uses the `@volatile` tag.</span></span>

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a><span data-ttu-id="d6783-229">类型</span><span class="sxs-lookup"><span data-stu-id="d6783-229">Types</span></span>

<span data-ttu-id="d6783-230">通过指定参数类型，Excel 会在调用函数之前将值转换为该类型。</span><span class="sxs-lookup"><span data-stu-id="d6783-230">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="d6783-231">如果类型为 `any`，则不会执行任何转换。</span><span class="sxs-lookup"><span data-stu-id="d6783-231">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="d6783-232">值类型</span><span class="sxs-lookup"><span data-stu-id="d6783-232">Value types</span></span>

<span data-ttu-id="d6783-233">可以使用以下类型之一表示单个值：`boolean`、`number`、`string`。</span><span class="sxs-lookup"><span data-stu-id="d6783-233">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="d6783-234">矩阵类型</span><span class="sxs-lookup"><span data-stu-id="d6783-234">Matrix type</span></span>

<span data-ttu-id="d6783-235">使用二维数组类型将参数或返回值变为值的矩阵。</span><span class="sxs-lookup"><span data-stu-id="d6783-235">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="d6783-236">例如，类型 `number[][]` 指示数字矩阵， `string[][]` 并指示字符串矩阵。</span><span class="sxs-lookup"><span data-stu-id="d6783-236">For example, the type `number[][]` indicates a matrix of numbers and `string[][]` indicates a matrix of strings.</span></span>

### <a name="error-type"></a><span data-ttu-id="d6783-237">错误类型</span><span class="sxs-lookup"><span data-stu-id="d6783-237">Error type</span></span>

<span data-ttu-id="d6783-238">非流式处理函数可以通过返回错误类型来指示错误。</span><span class="sxs-lookup"><span data-stu-id="d6783-238">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="d6783-239">流式处理函数可以通过使用错误类型调用 `setResult()` 来指示错误。</span><span class="sxs-lookup"><span data-stu-id="d6783-239">A streaming function can indicate an error by calling `setResult()` with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="d6783-240">Promise</span><span class="sxs-lookup"><span data-stu-id="d6783-240">Promise</span></span>

<span data-ttu-id="d6783-241">自定义函数可以返回一个承诺，该承诺在承诺实现时提供值。</span><span class="sxs-lookup"><span data-stu-id="d6783-241">A custom function can return a promise that provides the value when the promise is resolved.</span></span> <span data-ttu-id="d6783-242">如果承诺被拒绝，则自定义函数将引发错误。</span><span class="sxs-lookup"><span data-stu-id="d6783-242">If the promise is rejected, then the custom function will throw an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="d6783-243">其他类型</span><span class="sxs-lookup"><span data-stu-id="d6783-243">Other types</span></span>

<span data-ttu-id="d6783-244">任何其他类型都将被视为错误。</span><span class="sxs-lookup"><span data-stu-id="d6783-244">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="d6783-245">后续步骤</span><span class="sxs-lookup"><span data-stu-id="d6783-245">Next steps</span></span>

<span data-ttu-id="d6783-246">了解[自定义函数的命名约定](custom-functions-naming.md)。</span><span class="sxs-lookup"><span data-stu-id="d6783-246">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="d6783-247">或者，了解如何[本地化函数](custom-functions-localize.md)，这需要你[手动编写 JSON 文件](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="d6783-247">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d6783-248">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d6783-248">See also</span></span>

* [<span data-ttu-id="d6783-249">手动为自定义函数创建 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="d6783-249">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d6783-250">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="d6783-250">Create custom functions in Excel</span></span>](custom-functions-overview.md)

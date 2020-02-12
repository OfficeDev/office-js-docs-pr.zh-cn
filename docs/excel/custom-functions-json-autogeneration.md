---
ms.date: 07/15/2019
description: 使用 JSDoc 标记动态创建自定义函数 JSON 元数据。
title: 为自定义函数自动生成 JSON 元数据
localization_priority: Normal
ms.openlocfilehash: 16db164b061840bd8dfb44124f956da773ad9d2b
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950843"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="3f26a-103">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="3f26a-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="3f26a-104">在使用 JavaScript 或 TypeScript 编写 Excel 自定义函数时，使用 [JSDoc 标记](https://jsdoc.app/)提供有关自定义函数的额外信息。</span><span class="sxs-lookup"><span data-stu-id="3f26a-104">When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function.</span></span> <span data-ttu-id="3f26a-105">然后在生成时使用 JSDoc 标记创建 [JSON 元数据文件](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="3f26a-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="3f26a-106">使用 JSDoc 标记使您免除手动编辑 JSON 元数据文件的工作。</span><span class="sxs-lookup"><span data-stu-id="3f26a-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="3f26a-107">为 JavaScript 或 TypeScript 函数添加代码注释中的 `@customfunction` 标记以将其标记为自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="3f26a-108">可以使用 JavaScript 中的 [@param](#param) 标记或从 TypeScript 中的[函数类型](https://www.typescriptlang.org/docs/handbook/functions.html)提供函数参数类型。</span><span class="sxs-lookup"><span data-stu-id="3f26a-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="3f26a-109">有关详细信息，请参阅 [@param](#param) 标记和[类型](#types)部分。</span><span class="sxs-lookup"><span data-stu-id="3f26a-109">For more information, see the [@param](#param) tag and [Types](#types) sections.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="3f26a-110">为函数添加说明</span><span class="sxs-lookup"><span data-stu-id="3f26a-110">Adding a description to a function</span></span>

<span data-ttu-id="3f26a-111">当用户需要帮助来了解自定义函数的功能时，将向用户显示用作帮助文本的说明。</span><span class="sxs-lookup"><span data-stu-id="3f26a-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="3f26a-112">说明不需要任何特定标记。</span><span class="sxs-lookup"><span data-stu-id="3f26a-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="3f26a-113">只需在 JSDoc 注释中输入简短的文本说明即可。</span><span class="sxs-lookup"><span data-stu-id="3f26a-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="3f26a-114">一般来说，说明位于 JSDoc 注释部分的开头，但无论位于何处，它都有用。</span><span class="sxs-lookup"><span data-stu-id="3f26a-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="3f26a-115">若要查看内置函数说明的示例，请打开 Excel，转到“**公式**”选项卡，然后选择“**插入函数**”。</span><span class="sxs-lookup"><span data-stu-id="3f26a-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="3f26a-116">然后，你可以浏览所有函数说明，还可以查看列出的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="3f26a-117">在以下示例中，短语“计算球体的体积”</span><span class="sxs-lookup"><span data-stu-id="3f26a-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="3f26a-118">就是自定义函数的相关说明。</span><span class="sxs-lookup"><span data-stu-id="3f26a-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="3f26a-119">JSDoc 标记</span><span class="sxs-lookup"><span data-stu-id="3f26a-119">JSDoc Tags</span></span>
<span data-ttu-id="3f26a-120">Excel 自定义函数支持以下 JSDoc 标记：</span><span class="sxs-lookup"><span data-stu-id="3f26a-120">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="3f26a-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="3f26a-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="3f26a-122">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="3f26a-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="3f26a-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="3f26a-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="3f26a-124">[@param](#param) _{type}_ name description</span><span class="sxs-lookup"><span data-stu-id="3f26a-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="3f26a-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="3f26a-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="3f26a-126">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="3f26a-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="3f26a-127">@streaming</span><span class="sxs-lookup"><span data-stu-id="3f26a-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="3f26a-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="3f26a-128">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="3f26a-129">@cancelable</span><span class="sxs-lookup"><span data-stu-id="3f26a-129">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="3f26a-130">表示自定义函数希望在取消函数时执行操作。</span><span class="sxs-lookup"><span data-stu-id="3f26a-130">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="3f26a-131">最后一个函数参数的类型必须是 `CustomFunctions.CancelableInvocation`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="3f26a-132">该函数可以将函数分配给 `oncanceled` 属性来表示在取消函数时要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="3f26a-132">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="3f26a-133">如果最后一个函数参数的类型为 `CustomFunctions.CancelableInvocation`，则即使标记不存在，也会被视为 `@cancelable`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.</span></span>

<span data-ttu-id="3f26a-134">函数不能同时具有 `@cancelable` 和 `@streaming` 标记。</span><span class="sxs-lookup"><span data-stu-id="3f26a-134">A function can't have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="3f26a-135">@customfunction</span><span class="sxs-lookup"><span data-stu-id="3f26a-135">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="3f26a-136">语法：@customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="3f26a-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="3f26a-137">指定此标记以将 JavaScript/TypeScript 函数视为 Excel 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-137">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="3f26a-138">需要此标记才能创建自定义函数的元数据。</span><span class="sxs-lookup"><span data-stu-id="3f26a-138">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="3f26a-139">下面的示例显示了声明自定义函数的最简单方法。</span><span class="sxs-lookup"><span data-stu-id="3f26a-139">The following example shows the simplest way to declare a custom function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="3f26a-140">id</span><span class="sxs-lookup"><span data-stu-id="3f26a-140">id</span></span>

<span data-ttu-id="3f26a-141">`id` 是自定义函数的固定标识符。</span><span class="sxs-lookup"><span data-stu-id="3f26a-141">The `id` is an invariant identifier for the custom function.</span></span>

* <span data-ttu-id="3f26a-142">如果未提供 `id`，请将 JavaScript/TypeScript 函数名称转换为大写并删除禁用字符。</span><span class="sxs-lookup"><span data-stu-id="3f26a-142">If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.</span></span>
* <span data-ttu-id="3f26a-143">`id` 对于所有自定义函数必须是唯一的。</span><span class="sxs-lookup"><span data-stu-id="3f26a-143">The `id` must be unique for all custom functions.</span></span>
* <span data-ttu-id="3f26a-144">允许使用的字符限为：A-Z、a-z、0-9、下划线 (\_) 和句点 (.)。</span><span class="sxs-lookup"><span data-stu-id="3f26a-144">The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).</span></span>

<span data-ttu-id="3f26a-145">在下面的示例中，增量是函数的 `id` 和 `name`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-145">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="3f26a-146">name</span><span class="sxs-lookup"><span data-stu-id="3f26a-146">name</span></span>

<span data-ttu-id="3f26a-147">提供自定义函数的显示`name`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-147">Provides the display `name` for the custom function.</span></span>

* <span data-ttu-id="3f26a-148">如果未提供名称，则 id 还会用作名称。</span><span class="sxs-lookup"><span data-stu-id="3f26a-148">If name isn't provided, the id is also used as the name.</span></span>
* <span data-ttu-id="3f26a-149">允许使用的字符：字母 [Unicode 字母字符](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、句点 (.) 和下划线 (\_)。</span><span class="sxs-lookup"><span data-stu-id="3f26a-149">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="3f26a-150">必须以字母开头。</span><span class="sxs-lookup"><span data-stu-id="3f26a-150">Must start with a letter.</span></span>
* <span data-ttu-id="3f26a-151">最大长度为 128 个字符。</span><span class="sxs-lookup"><span data-stu-id="3f26a-151">Maximum length is 128 characters.</span></span>

<span data-ttu-id="3f26a-152">在下面的示例中，INC 是函数的 `id`，并且 `increment` 是 `name`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-152">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="3f26a-153">说明</span><span class="sxs-lookup"><span data-stu-id="3f26a-153">description</span></span>

<span data-ttu-id="3f26a-154">说明不需要任何特定标记。</span><span class="sxs-lookup"><span data-stu-id="3f26a-154">A description doesn't require any specific tag.</span></span> <span data-ttu-id="3f26a-155">通过在 JSDoc 注释中添加一个短语来描述函数的功能，为自定义函数添加说明。</span><span class="sxs-lookup"><span data-stu-id="3f26a-155">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="3f26a-156">默认情况下，JSDoc 注释部分中未标记的任何文本都是该函数的说明。</span><span class="sxs-lookup"><span data-stu-id="3f26a-156">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span> <span data-ttu-id="3f26a-157">当 Excel 中的用户进入该函数时，将向其显示相关说明。</span><span class="sxs-lookup"><span data-stu-id="3f26a-157">The description appears to users in Excel as they are entering the function.</span></span> <span data-ttu-id="3f26a-158">在以下示例中，短语“对两个数字求和的函数”是 id 属性为 `ADD` 的自定义函数的相关说明。</span><span class="sxs-lookup"><span data-stu-id="3f26a-158">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

<span data-ttu-id="3f26a-159">在下面的示例中，ADD 是函数的 `id` 和 `name`，并且提供了说明。</span><span class="sxs-lookup"><span data-stu-id="3f26a-159">In the following example, ADD is the `id` and `name` of the function and a description is given.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
### <a name="helpurl"></a><span data-ttu-id="3f26a-160">@helpurl</span><span class="sxs-lookup"><span data-stu-id="3f26a-160">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="3f26a-161">语法：@helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="3f26a-161">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="3f26a-162">提供的 _url_ 显示在 Excel 中。</span><span class="sxs-lookup"><span data-stu-id="3f26a-162">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="3f26a-163">在下面的示例中，`helpurl` 是 www.contoso.com/weatherhelp。</span><span class="sxs-lookup"><span data-stu-id="3f26a-163">In the following example, the `helpurl` is www.contoso.com/weatherhelp.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
### <a name="param"></a><span data-ttu-id="3f26a-164">@param</span><span class="sxs-lookup"><span data-stu-id="3f26a-164">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="3f26a-165">JavaScript</span><span class="sxs-lookup"><span data-stu-id="3f26a-165">JavaScript</span></span>

<span data-ttu-id="3f26a-166">JavaScript 语法：@param {type} name _description_</span><span class="sxs-lookup"><span data-stu-id="3f26a-166">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="3f26a-167">`{type}` 应在大括号内指定类型信息。</span><span class="sxs-lookup"><span data-stu-id="3f26a-167">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="3f26a-168">有关可能使用的类型的详细信息，请参阅[类型](#types)部分。</span><span class="sxs-lookup"><span data-stu-id="3f26a-168">See the [Types](#types) section for more information about the types which may be used.</span></span> <span data-ttu-id="3f26a-169">可选：如果未指定，则使用类型 `any`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-169">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="3f26a-170">`name` 指定 @param 标记适用于哪个参数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-170">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="3f26a-171">必需。</span><span class="sxs-lookup"><span data-stu-id="3f26a-171">Required.</span></span>
* <span data-ttu-id="3f26a-172">`description` 为函数参数提供显示在 Excel 中的说明。</span><span class="sxs-lookup"><span data-stu-id="3f26a-172">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="3f26a-173">可选。</span><span class="sxs-lookup"><span data-stu-id="3f26a-173">Optional.</span></span>

<span data-ttu-id="3f26a-174">若要将自定义函数参数表示为可选，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="3f26a-174">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="3f26a-175">为参数名称加上方括号。</span><span class="sxs-lookup"><span data-stu-id="3f26a-175">Put square brackets around the parameter name.</span></span> <span data-ttu-id="3f26a-176">例如：`@param {string} [text] Optional text`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-176">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="3f26a-177">可选参数的默认值为 `null`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-177">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="3f26a-178">下面的示例显示了 ADD 函数，该函数将两个或三个数字相加，第三个数字作为可选参数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-178">The following example shows a ADD function which adds two or three numbers, with the third number as an optional parameter.</span></span>

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

#### <a name="typescript"></a><span data-ttu-id="3f26a-179">TypeScript</span><span class="sxs-lookup"><span data-stu-id="3f26a-179">TypeScript</span></span>

<span data-ttu-id="3f26a-180">TypeScript 语法：@param name _description_</span><span class="sxs-lookup"><span data-stu-id="3f26a-180">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="3f26a-181">`name` 指定 @param 标记适用于哪个参数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-181">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="3f26a-182">必需。</span><span class="sxs-lookup"><span data-stu-id="3f26a-182">Required.</span></span>
* <span data-ttu-id="3f26a-183">`description` 为函数参数提供显示在 Excel 中的说明。</span><span class="sxs-lookup"><span data-stu-id="3f26a-183">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="3f26a-184">可选。</span><span class="sxs-lookup"><span data-stu-id="3f26a-184">Optional.</span></span>

<span data-ttu-id="3f26a-185">有关可能使用的函数参数类型的详细信息，请参阅[类型](#types)部分。</span><span class="sxs-lookup"><span data-stu-id="3f26a-185">See the [Types](#types) section for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="3f26a-186">若要将自定义函数参数表示为可选，请执行以下操作之一：</span><span class="sxs-lookup"><span data-stu-id="3f26a-186">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="3f26a-187">使用可选参数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-187">Use an optional parameter.</span></span> <span data-ttu-id="3f26a-188">例如：`function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="3f26a-188">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="3f26a-189">为该参数提供默认值。</span><span class="sxs-lookup"><span data-stu-id="3f26a-189">Give the parameter a default value.</span></span> <span data-ttu-id="3f26a-190">例如：`function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="3f26a-190">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="3f26a-191">有关 @param 的详细说明，请参阅：[JSDoc](https://jsdoc.app/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="3f26a-191">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="3f26a-192">可选参数的默认值为 `null`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-192">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="3f26a-193">下面的示例显示了将两个数字相加的 `add` 函数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-193">The following example shows the `add` function that adds two numbers.</span></span>

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

---
### <a name="requiresaddress"></a><span data-ttu-id="3f26a-194">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="3f26a-194">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="3f26a-195">表示应提供计算函数所在的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="3f26a-195">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="3f26a-196">最后一个函数参数的类型必须是 `CustomFunctions.Invocation` 或派生类型。</span><span class="sxs-lookup"><span data-stu-id="3f26a-196">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="3f26a-197">调用函数时，`address` 属性将包含地址。</span><span class="sxs-lookup"><span data-stu-id="3f26a-197">When the function is called, the `address` property will contain the address.</span></span> <span data-ttu-id="3f26a-198">有关使用 `@requiresAddress` 标记的函数示例，请参阅[寻址单元格的上下文参数](custom-functions-parameter-options.md#addressing-cells-context-parameter)。</span><span class="sxs-lookup"><span data-stu-id="3f26a-198">For an example of a function that uses the `@requiresAddress` tag, see [Addressing cell's context parameter](custom-functions-parameter-options.md#addressing-cells-context-parameter).</span></span>

---
### <a name="returns"></a><span data-ttu-id="3f26a-199">@returns</span><span class="sxs-lookup"><span data-stu-id="3f26a-199">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="3f26a-200">语法：@returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="3f26a-200">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="3f26a-201">提供返回值的类型。</span><span class="sxs-lookup"><span data-stu-id="3f26a-201">Provides the type for the return value.</span></span>

<span data-ttu-id="3f26a-202">如果省略 `{type}`，则将使用 TypeScript 类型信息。</span><span class="sxs-lookup"><span data-stu-id="3f26a-202">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="3f26a-203">如果没有类型信息，则类型将为 `any`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-203">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="3f26a-204">下面的示例显示了使用 `@returns` 标记的 `add` 函数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-204">The following example shows the `add` function that uses the `@returns` tag.</span></span>

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

---
### <a name="streaming"></a><span data-ttu-id="3f26a-205">@streaming</span><span class="sxs-lookup"><span data-stu-id="3f26a-205">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="3f26a-206">用于表示自定义函数是一个流式处理函数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-206">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="3f26a-207">最后一个参数的类型应为 `CustomFunctions.StreamingInvocation<ResultType>`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-207">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="3f26a-208">该函数应返回 `void`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-208">The function should return `void`.</span></span>

<span data-ttu-id="3f26a-209">流式处理函数不直接返回值，而是应该使用最后一个参数调用 `setResult(result: ResultType)`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-209">Streaming functions don't return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="3f26a-210">由流式处理函数引发的异常将被忽略。</span><span class="sxs-lookup"><span data-stu-id="3f26a-210">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="3f26a-211">`setResult()` 可能称为“错误”，以指示错误结果。</span><span class="sxs-lookup"><span data-stu-id="3f26a-211">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="3f26a-212">有关流式处理函数的示例和更多信息，请参阅[生成流式处理函数](./custom-functions-web-reqs.md#make-a-streaming-function)。</span><span class="sxs-lookup"><span data-stu-id="3f26a-212">For an example of a streaming function and more information, see [Make a streaming function](./custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="3f26a-213">流式处理函数不能标记为 [@volatile](#volatile)。</span><span class="sxs-lookup"><span data-stu-id="3f26a-213">Streaming functions can't be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="3f26a-214">@volatile</span><span class="sxs-lookup"><span data-stu-id="3f26a-214">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="3f26a-215">可变函数是指其结果不断变化的函数，即使不采用任何参数或参数未发生更改都是如此。</span><span class="sxs-lookup"><span data-stu-id="3f26a-215">A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed.</span></span> <span data-ttu-id="3f26a-216">Excel 在每次完成计算后，都会重新计算包含可变函数和所有依赖项的单元格。</span><span class="sxs-lookup"><span data-stu-id="3f26a-216">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="3f26a-217">因此，过于依赖可变函数会使重新计算时间变慢，请谨慎使用。</span><span class="sxs-lookup"><span data-stu-id="3f26a-217">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="3f26a-218">流式处理函数不能为可变函数。</span><span class="sxs-lookup"><span data-stu-id="3f26a-218">Streaming functions can't be volatile.</span></span>

<span data-ttu-id="3f26a-219">以下函数是可变函数并使用 `@volatile` 标记。</span><span class="sxs-lookup"><span data-stu-id="3f26a-219">The following function is volatile and uses the `@volatile` tag.</span></span>

```js
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a><span data-ttu-id="3f26a-220">类型</span><span class="sxs-lookup"><span data-stu-id="3f26a-220">Types</span></span>

<span data-ttu-id="3f26a-221">通过指定参数类型，Excel 会在调用函数之前将值转换为该类型。</span><span class="sxs-lookup"><span data-stu-id="3f26a-221">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="3f26a-222">如果类型为 `any`，则不会执行任何转换。</span><span class="sxs-lookup"><span data-stu-id="3f26a-222">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="3f26a-223">值类型</span><span class="sxs-lookup"><span data-stu-id="3f26a-223">Value types</span></span>

<span data-ttu-id="3f26a-224">可以使用以下类型之一表示单个值：`boolean`、`number`、`string`。</span><span class="sxs-lookup"><span data-stu-id="3f26a-224">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="3f26a-225">矩阵类型</span><span class="sxs-lookup"><span data-stu-id="3f26a-225">Matrix type</span></span>

<span data-ttu-id="3f26a-226">使用二维数组类型将参数或返回值变为值的矩阵。</span><span class="sxs-lookup"><span data-stu-id="3f26a-226">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="3f26a-227">例如，类型 `number[][]` 表示数字的矩阵。</span><span class="sxs-lookup"><span data-stu-id="3f26a-227">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="3f26a-228">`string[][]` 表示字符串的矩阵。</span><span class="sxs-lookup"><span data-stu-id="3f26a-228">`string[][]` indicates a matrix of strings.</span></span>

### <a name="error-type"></a><span data-ttu-id="3f26a-229">错误类型</span><span class="sxs-lookup"><span data-stu-id="3f26a-229">Error type</span></span>

<span data-ttu-id="3f26a-230">非流式处理函数可以通过返回错误类型来指示错误。</span><span class="sxs-lookup"><span data-stu-id="3f26a-230">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="3f26a-231">流式处理函数可以通过使用错误类型调用 `setResult()` 来指示错误。</span><span class="sxs-lookup"><span data-stu-id="3f26a-231">A streaming function can indicate an error by calling `setResult()` with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="3f26a-232">Promise</span><span class="sxs-lookup"><span data-stu-id="3f26a-232">Promise</span></span>

<span data-ttu-id="3f26a-233">函数可以返回 Promise，将在解析 promise 后提供值。</span><span class="sxs-lookup"><span data-stu-id="3f26a-233">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="3f26a-234">如果 promise 被拒绝，则会出现错误。</span><span class="sxs-lookup"><span data-stu-id="3f26a-234">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="3f26a-235">其他类型</span><span class="sxs-lookup"><span data-stu-id="3f26a-235">Other types</span></span>

<span data-ttu-id="3f26a-236">任何其他类型都将被视为错误。</span><span class="sxs-lookup"><span data-stu-id="3f26a-236">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="3f26a-237">后续步骤</span><span class="sxs-lookup"><span data-stu-id="3f26a-237">Next steps</span></span>
<span data-ttu-id="3f26a-238">了解[自定义函数的命名约定](custom-functions-naming.md)。</span><span class="sxs-lookup"><span data-stu-id="3f26a-238">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="3f26a-239">或者，了解如何[本地化函数](custom-functions-localize.md)，这需要你[手动编写 JSON 文件](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="3f26a-239">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3f26a-240">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3f26a-240">See also</span></span>

* [<span data-ttu-id="3f26a-241">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="3f26a-241">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="3f26a-242">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="3f26a-242">Create custom functions in Excel</span></span>](custom-functions-overview.md)

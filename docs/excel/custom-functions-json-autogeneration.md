---
ms.date: 04/25/2019
description: 使用 JSDOC 标记动态创建自定义函数 JSON 元数据。
title: 创建自定义函数的 JSON 元数据（预览）
localization_priority: Priority
ms.openlocfilehash: f13c9b2ff0adfd1bda493433761dc74cb9ea37dc
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353270"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="53120-103">创建自定义函数的 JSON 元数据（预览）</span><span class="sxs-lookup"><span data-stu-id="53120-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="53120-104">在 JavaScript 或 TypeScript 中写入 Excel 自定义函数时，使用 JSDoc 标记提供有关自定义函数的额外信息。</span><span class="sxs-lookup"><span data-stu-id="53120-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="53120-105">然后在生成时使用 JSDoc 标记创建 [JSON 元数据文件](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="53120-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="53120-106">使用 JSDoc 标记使您免除手动编辑 JSON 元数据文件的工作。</span><span class="sxs-lookup"><span data-stu-id="53120-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="53120-107">为 JavaScript 或 TypeScript 函数添加代码注释中的 `@customfunction` 标记以将其标记为自定义函数。</span><span class="sxs-lookup"><span data-stu-id="53120-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="53120-108">可以使用 JavaScript 中的 [@param](#param) 标记或从 TypeScript 中的[函数类型](https://www.typescriptlang.org/docs/handbook/functions.html)提供函数参数类型。</span><span class="sxs-lookup"><span data-stu-id="53120-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="53120-109">有关详细信息，请参阅 [@param](#param) 标记和[类型](#types)部分。</span><span class="sxs-lookup"><span data-stu-id="53120-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="53120-110">JSDoc 标记</span><span class="sxs-lookup"><span data-stu-id="53120-110">JSDoc Tags</span></span>
<span data-ttu-id="53120-111">Excel 自定义函数支持以下 JSDoc 标记：</span><span class="sxs-lookup"><span data-stu-id="53120-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="53120-112">@cancelable</span><span class="sxs-lookup"><span data-stu-id="53120-112">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="53120-113">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="53120-113">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="53120-114">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="53120-114">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="53120-115">[@param](#param) _{type}_ name description</span><span class="sxs-lookup"><span data-stu-id="53120-115">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="53120-116">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="53120-116">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="53120-117">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="53120-117">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="53120-118">@streaming</span><span class="sxs-lookup"><span data-stu-id="53120-118">@streaming</span></span>](#streaming)
* [<span data-ttu-id="53120-119">@volatile</span><span class="sxs-lookup"><span data-stu-id="53120-119">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="53120-120">@cancelable</span><span class="sxs-lookup"><span data-stu-id="53120-120">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="53120-121">表示自定义函数希望在取消函数时执行操作。</span><span class="sxs-lookup"><span data-stu-id="53120-121">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="53120-122">最后一个函数参数的类型必须是 `CustomFunctions.CancelableInvocation`。</span><span class="sxs-lookup"><span data-stu-id="53120-122">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="53120-123">该函数可以将函数分配给 `oncanceled` 属性来表示在取消函数时要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="53120-123">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="53120-124">如果最后一个函数参数的类型为 `CustomFunctions.CancelableInvocation`，则即使标记不存在，也会被视为 `@cancelable`。</span><span class="sxs-lookup"><span data-stu-id="53120-124">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="53120-125">函数不能同时具有 `@cancelable` 和 `@streaming` 标记。</span><span class="sxs-lookup"><span data-stu-id="53120-125">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="53120-126">@customfunction</span><span class="sxs-lookup"><span data-stu-id="53120-126">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="53120-127">语法：@customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="53120-127">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="53120-128">指定此标记以将 JavaScript/TypeScript 函数视为 Excel 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="53120-128">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="53120-129">需要此标记才能创建自定义函数的元数据。</span><span class="sxs-lookup"><span data-stu-id="53120-129">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="53120-130">还应调用 `CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="53120-130">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="53120-131">id</span><span class="sxs-lookup"><span data-stu-id="53120-131">id</span></span>

<span data-ttu-id="53120-132">id 用作存储在文档中的自定义函数的固定标识符。</span><span class="sxs-lookup"><span data-stu-id="53120-132">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="53120-133">不得更改。</span><span class="sxs-lookup"><span data-stu-id="53120-133">It should not change.</span></span>

* <span data-ttu-id="53120-134">如果未提供 id，JavaScript/TypeScript 函数名称将转换为大写形式，并删除不允许使用的字符。</span><span class="sxs-lookup"><span data-stu-id="53120-134">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="53120-135">id 对于所有自定义函数必须是唯一的。</span><span class="sxs-lookup"><span data-stu-id="53120-135">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="53120-136">允许使用的字符仅限于：A-Z、a-z、0-9 和句点 (.)。</span><span class="sxs-lookup"><span data-stu-id="53120-136">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="53120-137">name</span><span class="sxs-lookup"><span data-stu-id="53120-137">name</span></span>

<span data-ttu-id="53120-138">提供自定义函数的显示名称。</span><span class="sxs-lookup"><span data-stu-id="53120-138">Provides the display name for the custom function.</span></span>

* <span data-ttu-id="53120-139">如果未提供名称，则 id 还会用作名称。</span><span class="sxs-lookup"><span data-stu-id="53120-139">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="53120-140">允许使用的字符：字母 [Unicode 字母字符](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、句点 (.) 和下划线 (\_)。</span><span class="sxs-lookup"><span data-stu-id="53120-140">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="53120-141">必须以字母开头。</span><span class="sxs-lookup"><span data-stu-id="53120-141">Must start with a letter.</span></span>
* <span data-ttu-id="53120-142">最大长度为 128 个字符。</span><span class="sxs-lookup"><span data-stu-id="53120-142">Maximum length is 128 characters.</span></span>

---
### <a name="helpurl"></a><span data-ttu-id="53120-143">@helpurl</span><span class="sxs-lookup"><span data-stu-id="53120-143">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="53120-144">语法：@helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="53120-144">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="53120-145">提供的 _url_ 显示在 Excel 中。</span><span class="sxs-lookup"><span data-stu-id="53120-145">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="53120-146">@param</span><span class="sxs-lookup"><span data-stu-id="53120-146">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="53120-147">JavaScript</span><span class="sxs-lookup"><span data-stu-id="53120-147">JavaScript</span></span>

<span data-ttu-id="53120-148">JavaScript 语法：@param {type} name _description_</span><span class="sxs-lookup"><span data-stu-id="53120-148">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="53120-149">`{type}` 应在大括号内指定类型信息。</span><span class="sxs-lookup"><span data-stu-id="53120-149">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="53120-150">有关可能使用的类型的详细信息，请参阅[类型](##types)。</span><span class="sxs-lookup"><span data-stu-id="53120-150">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="53120-151">可选：如果未指定，则使用类型 `any`。</span><span class="sxs-lookup"><span data-stu-id="53120-151">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="53120-152">`name` 指定 @param 标记适用于哪个参数。</span><span class="sxs-lookup"><span data-stu-id="53120-152">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="53120-153">必需。</span><span class="sxs-lookup"><span data-stu-id="53120-153">Required.</span></span>
* <span data-ttu-id="53120-154">`description` 为函数参数提供显示在 Excel 中的说明。</span><span class="sxs-lookup"><span data-stu-id="53120-154">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="53120-155">可选。</span><span class="sxs-lookup"><span data-stu-id="53120-155">Optional.</span></span>

<span data-ttu-id="53120-156">若要将自定义函数参数表示为可选，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="53120-156">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="53120-157">为参数名称加上方括号。</span><span class="sxs-lookup"><span data-stu-id="53120-157">Put square brackets around the parameter name.</span></span> <span data-ttu-id="53120-158">例如：`@param {string} [text] Optional text`。</span><span class="sxs-lookup"><span data-stu-id="53120-158">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="53120-159">可选参数的默认值为 `null`。</span><span class="sxs-lookup"><span data-stu-id="53120-159">The default value for optional parameters is `null`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="53120-160">TypeScript</span><span class="sxs-lookup"><span data-stu-id="53120-160">TypeScript</span></span>

<span data-ttu-id="53120-161">TypeScript 语法：@param name _description_</span><span class="sxs-lookup"><span data-stu-id="53120-161">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="53120-162">`name` 指定 @param 标记适用于哪个参数。</span><span class="sxs-lookup"><span data-stu-id="53120-162">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="53120-163">必需。</span><span class="sxs-lookup"><span data-stu-id="53120-163">Required.</span></span>
* <span data-ttu-id="53120-164">`description` 为函数参数提供显示在 Excel 中的说明。</span><span class="sxs-lookup"><span data-stu-id="53120-164">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="53120-165">可选。</span><span class="sxs-lookup"><span data-stu-id="53120-165">Optional.</span></span>

<span data-ttu-id="53120-166">有关可能使用的函数参数类型的详细信息，请参阅[类型](##types)。</span><span class="sxs-lookup"><span data-stu-id="53120-166">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="53120-167">若要将自定义函数参数表示为可选，请执行以下操作之一：</span><span class="sxs-lookup"><span data-stu-id="53120-167">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="53120-168">使用可选参数。</span><span class="sxs-lookup"><span data-stu-id="53120-168">Use an optional parameter.</span></span> <span data-ttu-id="53120-169">例如：`function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="53120-169">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="53120-170">为该参数提供默认值。</span><span class="sxs-lookup"><span data-stu-id="53120-170">Give the parameter a default value.</span></span> <span data-ttu-id="53120-171">例如：`function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="53120-171">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="53120-172">有关 @param 的详细说明，请参阅：[JSDoc](http://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="53120-172">For detailed description of the @param see: [JSDoc](http://usejsdoc.org/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="53120-173">可选参数的默认值为 `null`。</span><span class="sxs-lookup"><span data-stu-id="53120-173">The default value for optional parameters is `null`.</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="53120-174">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="53120-174">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="53120-175">表示应提供计算函数所在的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="53120-175">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="53120-176">最后一个函数参数的类型必须是 `CustomFunctions.Invocation` 或派生类型。</span><span class="sxs-lookup"><span data-stu-id="53120-176">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="53120-177">调用函数时，`address` 属性将包含地址。</span><span class="sxs-lookup"><span data-stu-id="53120-177">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="53120-178">@returns</span><span class="sxs-lookup"><span data-stu-id="53120-178">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="53120-179">语法：@returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="53120-179">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="53120-180">提供返回值的类型。</span><span class="sxs-lookup"><span data-stu-id="53120-180">Provides the type for the return value.</span></span>

<span data-ttu-id="53120-181">如果省略 `{type}`，则将使用 TypeScript 类型信息。</span><span class="sxs-lookup"><span data-stu-id="53120-181">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="53120-182">如果没有类型信息，则类型将为 `any`。</span><span class="sxs-lookup"><span data-stu-id="53120-182">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="53120-183">@streaming</span><span class="sxs-lookup"><span data-stu-id="53120-183">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="53120-184">用于表示自定义函数是一个流式处理函数。</span><span class="sxs-lookup"><span data-stu-id="53120-184">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="53120-185">最后一个参数的类型应为 `CustomFunctions.StreamingInvocation<ResultType>`。</span><span class="sxs-lookup"><span data-stu-id="53120-185">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="53120-186">该函数应返回 `void`。</span><span class="sxs-lookup"><span data-stu-id="53120-186">The function should return `void`.</span></span>

<span data-ttu-id="53120-187">流式处理函数不直接返回值，而是应该使用最后一个参数调用 `setResult(result: ResultType)`。</span><span class="sxs-lookup"><span data-stu-id="53120-187">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="53120-188">由流式处理函数引发的异常将被忽略。</span><span class="sxs-lookup"><span data-stu-id="53120-188">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="53120-189">`setResult()` 可能称为“错误”，以指示错误结果。</span><span class="sxs-lookup"><span data-stu-id="53120-189">`setResult()` may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="53120-190">流式处理函数不能标记为 [@volatile](#volatile)。</span><span class="sxs-lookup"><span data-stu-id="53120-190">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="53120-191">@volatile</span><span class="sxs-lookup"><span data-stu-id="53120-191">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="53120-192">可变函数是其结果不能假定为即使不采用任何参数或参数未发生更改也始终保持不变的函数。</span><span class="sxs-lookup"><span data-stu-id="53120-192">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="53120-193">Excel 在每次完成计算后，都会重新计算包含可变函数和所有依赖项的单元格。</span><span class="sxs-lookup"><span data-stu-id="53120-193">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="53120-194">因此，过于依赖可变函数会使重新计算时间变慢，请谨慎使用。</span><span class="sxs-lookup"><span data-stu-id="53120-194">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="53120-195">流式处理函数不能为可变函数。</span><span class="sxs-lookup"><span data-stu-id="53120-195">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="53120-196">类型</span><span class="sxs-lookup"><span data-stu-id="53120-196">Types</span></span>

<span data-ttu-id="53120-197">通过指定参数类型，Excel 会在调用函数之前将值转换为该类型。</span><span class="sxs-lookup"><span data-stu-id="53120-197">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="53120-198">如果类型为 `any`，则不会执行任何转换。</span><span class="sxs-lookup"><span data-stu-id="53120-198">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="53120-199">值类型</span><span class="sxs-lookup"><span data-stu-id="53120-199">Value types</span></span>

<span data-ttu-id="53120-200">可以使用以下类型之一表示单个值：`boolean`、`number`、`string`。</span><span class="sxs-lookup"><span data-stu-id="53120-200">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="53120-201">矩阵类型</span><span class="sxs-lookup"><span data-stu-id="53120-201">Matrix type</span></span>

<span data-ttu-id="53120-202">使用二维数组类型将参数或返回值变为值的矩阵。</span><span class="sxs-lookup"><span data-stu-id="53120-202">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="53120-203">例如，类型 `number[][]` 表示数字的矩阵。</span><span class="sxs-lookup"><span data-stu-id="53120-203">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="53120-204">`string[][]` 表示字符串的矩阵。</span><span class="sxs-lookup"><span data-stu-id="53120-204">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="53120-205">错误类型</span><span class="sxs-lookup"><span data-stu-id="53120-205">Error type</span></span>

<span data-ttu-id="53120-206">非流式处理函数可以通过返回错误类型来指示错误。</span><span class="sxs-lookup"><span data-stu-id="53120-206">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="53120-207">流式处理函数可以通过使用错误类型调用 setResult() 来指示错误。</span><span class="sxs-lookup"><span data-stu-id="53120-207">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="53120-208">Promise</span><span class="sxs-lookup"><span data-stu-id="53120-208">Promise</span></span>

<span data-ttu-id="53120-209">函数可以返回 Promise，将在解析 promise 后提供值。</span><span class="sxs-lookup"><span data-stu-id="53120-209">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="53120-210">如果 promise 被拒绝，则会出现错误。</span><span class="sxs-lookup"><span data-stu-id="53120-210">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="53120-211">其他类型</span><span class="sxs-lookup"><span data-stu-id="53120-211">Other types</span></span>

<span data-ttu-id="53120-212">任何其他类型都将被视为错误。</span><span class="sxs-lookup"><span data-stu-id="53120-212">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="53120-213">另请参阅</span><span class="sxs-lookup"><span data-stu-id="53120-213">See also</span></span>

* [<span data-ttu-id="53120-214">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="53120-214">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="53120-215">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="53120-215">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="53120-216">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="53120-216">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="53120-217">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="53120-217">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="53120-218">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="53120-218">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="53120-219">自定义函数调试</span><span class="sxs-lookup"><span data-stu-id="53120-219">Custom functions debugging</span></span>](custom-functions-debugging.md)

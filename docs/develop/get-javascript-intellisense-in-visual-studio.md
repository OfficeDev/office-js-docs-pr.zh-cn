---
title: 在 Visual Studio 2019 中获取 JavaScript IntelliSense
description: 了解如何使用 JSDoc 为 JavaScript IntelliSense、对象、参数和返回值创建属性。
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: 6135649ce80e496d5e195b0ddb0dcb64172d41f5
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076053"
---
# <a name="get-javascript-intellisense-in-visual-studio-2019"></a><span data-ttu-id="7e856-103">在 Visual Studio 2019 中获取 JavaScript IntelliSense</span><span class="sxs-lookup"><span data-stu-id="7e856-103">Get JavaScript IntelliSense in Visual Studio 2019</span></span>

<span data-ttu-id="7e856-p101">当使用 Visual Studio 2019 开发 Office 外接程序时，可以使用 JSDoc 来启用 IntelliSense，以获取 JavaScript 变量、对象、参数和返回值。本文概述了 JSDoc 以及如何使用它在 Visual Studio 中创建 IntellSense。有关详细信息，请参阅 [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) 和 [JavaScript 中的 JSDoc 支持](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript)。</span><span class="sxs-lookup"><span data-stu-id="7e856-p101">When you use Visual Studio 2019 to develop Office Add-ins, you can use JSDoc to enable IntelliSense for your JavaScript variables, objects, parameters, and return values. This article provides an overview of JSDoc and how you can use it to create IntellSense in Visual Studio. For more details, see [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) and [JSDoc support in JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript).</span></span> 

## <a name="officejs-type-definitions"></a><span data-ttu-id="7e856-107">Office.js 类型定义</span><span class="sxs-lookup"><span data-stu-id="7e856-107">Office.js type definitions</span></span>

<span data-ttu-id="7e856-p102">需要向 Visual Studio 提供 Office.js 中的类型定义。为此，可以执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="7e856-p102">You need to provide the definitions of the types in Office.js to Visual Studio. To do this, you can:</span></span>

- <span data-ttu-id="7e856-p103">在名为 `\Office\1\` 的解决方案的某个文件夹中，创建 Office.js 文件的本地副本。在创建外接程序项目时，Visual Studio 中的 Office 外接程序项目模板会添加此本地副本。</span><span class="sxs-lookup"><span data-stu-id="7e856-p103">Have a local copy of the Office.js files in a folder in your solution named `\Office\1\`. The Office Add-in project templates in Visual Studio add this local copy when you create an add-in project.</span></span> 
- <span data-ttu-id="7e856-p104">通过将 tsconfig.json 文件添加到外接程序解决方案中的 Web 应用程序项目的根，来使用 Office.js 的在线版本。文件应当包括以下内容。</span><span class="sxs-lookup"><span data-stu-id="7e856-p104">Use an online version of Office.js by adding a tsconfig.json file to the root of the web application project in the add-in solution. The file should include the following content.</span></span>

    ```json
        {
            "compilerOptions": {
                "allowJs": true,            // These settings apply to JavaScript files also.
                "noEmit":  true             // Do not compile the JS (or TS) files in this project.
            },
            "exclude": [
                "node_modules",             // Don't include any JavaScript found under "node_modules".
                "Scripts/Office/1"          // Suppress loading all the JavaScript files from the Office NuGet package.
            ],
            "typeAcquisition": {
                "enable": true,             // Enable automatic fetching of type definitions for detected JavaScript libraries.
                "include": [ "office-js" ]  // Ensure that the "Office-js" type definition is fetched.
            }
        }
    ```

## <a name="jsdoc-syntax"></a><span data-ttu-id="7e856-114">JSDoc 语法</span><span class="sxs-lookup"><span data-stu-id="7e856-114">JSDoc syntax</span></span>

<span data-ttu-id="7e856-p105">基本技巧是在变量（或参数等）前面加上一个标识其数据类型的注释。这样，Visual Studio 中的 IntelliSense 可以推断其成员。示例如下。</span><span class="sxs-lookup"><span data-stu-id="7e856-p105">The basic technique is to precede the variable (or parameter, and so on) with a comment that identifies its data type. This allows IntelliSense in Visual Studio to infer its members. The following are examples.</span></span>

### <a name="variable"></a><span data-ttu-id="7e856-118">变量</span><span class="sxs-lookup"><span data-stu-id="7e856-118">Variable</span></span>

```js
/** @type {Excel.Range} */
var subsetRange;
```

![Screenshot showing excerpt of IntelliSense for 'subsetRange' variable.](../images/intellisense-vs17-var.png)

### <a name="parameter"></a><span data-ttu-id="7e856-120">参数</span><span class="sxs-lookup"><span data-stu-id="7e856-120">Parameter</span></span>

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```

![显示 JavaScript 示例 IntelliSense"paragraphs"参数中"paras"参数 (摘录的屏幕截图) 。](../images/intellisense-vs17-param.png)

### <a name="return-value"></a><span data-ttu-id="7e856-122">返回值</span><span class="sxs-lookup"><span data-stu-id="7e856-122">Return value</span></span>

```js
/** @returns {Word.Range} */
function myFunc() {

}
```

![Screenshot showing excerpt of IntelliSense for 'myFunc () ' return value.](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a><span data-ttu-id="7e856-124">复杂类型</span><span class="sxs-lookup"><span data-stu-id="7e856-124">Complex types</span></span>

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```

![Screenshot showing IntelliSense for complex type declaration of 'var myVar;' for example.](../images/intellisense-vs17-complex-type.png)

## <a name="see-also"></a><span data-ttu-id="7e856-126">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7e856-126">See also</span></span>

- [<span data-ttu-id="7e856-127">使用 Visual Studio 开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="7e856-127">Develop Office Add-ins with Visual Studio</span></span>](develop-add-ins-visual-studio.md)
- [<span data-ttu-id="7e856-128">在 Visual Studio 中调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="7e856-128">Debug Office Add-ins in Visual Studio</span></span>](debug-office-add-ins-in-visual-studio.md)

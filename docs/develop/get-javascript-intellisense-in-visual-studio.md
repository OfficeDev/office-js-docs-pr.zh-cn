---
title: 在 Visual Studio 2019 中获取 JavaScript IntelliSense
description: 了解如何使用 JSDoc 为 JavaScript 变量、对象、参数和返回值创建 IntelliSense。
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: 495e43994d78b1e01374e348e6d21d41d9611212
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131806"
---
# <a name="get-javascript-intellisense-in-visual-studio-2019"></a><span data-ttu-id="c0f84-103">在 Visual Studio 2019 中获取 JavaScript IntelliSense</span><span class="sxs-lookup"><span data-stu-id="c0f84-103">Get JavaScript IntelliSense in Visual Studio 2019</span></span>

<span data-ttu-id="c0f84-p101">当使用 Visual Studio 2019 开发 Office 外接程序时，可以使用 JSDoc 来启用 IntelliSense，以获取 JavaScript 变量、对象、参数和返回值。本文概述了 JSDoc 以及如何使用它在 Visual Studio 中创建 IntellSense。有关详细信息，请参阅 [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) 和 [JavaScript 中的 JSDoc 支持](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript)。</span><span class="sxs-lookup"><span data-stu-id="c0f84-p101">When you use Visual Studio 2019 to develop Office Add-ins, you can use JSDoc to enable IntelliSense for your JavaScript variables, objects, parameters, and return values. This article provides an overview of JSDoc and how you can use it to create IntellSense in Visual Studio. For more details, see [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) and [JSDoc support in JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript).</span></span> 

## <a name="officejs-type-definitions"></a><span data-ttu-id="c0f84-107">Office.js 类型定义</span><span class="sxs-lookup"><span data-stu-id="c0f84-107">Office.js type definitions</span></span>

<span data-ttu-id="c0f84-p102">需要向 Visual Studio 提供 Office.js 中的类型定义。为此，可以执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="c0f84-p102">You need to provide the definitions of the types in Office.js to Visual Studio. To do this, you can:</span></span>

- <span data-ttu-id="c0f84-p103">在名为 `\Office\1\` 的解决方案的某个文件夹中，创建 Office.js 文件的本地副本。在创建外接程序项目时，Visual Studio 中的 Office 外接程序项目模板会添加此本地副本。</span><span class="sxs-lookup"><span data-stu-id="c0f84-p103">Have a local copy of the Office.js files in a folder in your solution named `\Office\1\`. The Office Add-in project templates in Visual Studio add this local copy when you create an add-in project.</span></span> 
- <span data-ttu-id="c0f84-p104">通过将 tsconfig.json 文件添加到外接程序解决方案中的 Web 应用程序项目的根，来使用 Office.js 的在线版本。文件应当包括以下内容。</span><span class="sxs-lookup"><span data-stu-id="c0f84-p104">Use an online version of Office.js by adding a tsconfig.json file to the root of the web application project in the add-in solution. The file should include the following content.</span></span>

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

## <a name="jsdoc-syntax"></a><span data-ttu-id="c0f84-114">JSDoc 语法</span><span class="sxs-lookup"><span data-stu-id="c0f84-114">JSDoc syntax</span></span>

<span data-ttu-id="c0f84-p105">基本技巧是在变量（或参数等）前面加上一个标识其数据类型的注释。这样，Visual Studio 中的 IntelliSense 可以推断其成员。示例如下。</span><span class="sxs-lookup"><span data-stu-id="c0f84-p105">The basic technique is to precede the variable (or parameter, and so on) with a comment that identifies its data type. This allows IntelliSense in Visual Studio to infer its members. The following are examples.</span></span>

### <a name="variable"></a><span data-ttu-id="c0f84-118">变量</span><span class="sxs-lookup"><span data-stu-id="c0f84-118">Variable</span></span>

```js
/** @type {Excel.Range} */
var subsetRange;
```

![显示 "subsetRange" 变量的智能感知摘录的屏幕截图](../images/intellisense-vs17-var.png)

### <a name="parameter"></a><span data-ttu-id="c0f84-120">参数</span><span class="sxs-lookup"><span data-stu-id="c0f84-120">Parameter</span></span>

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```

![显示 JavaScript 示例中的 "paras" 参数的智能感知摘要的屏幕截图 ("段落" 参数) ](../images/intellisense-vs17-param.png)

### <a name="return-value"></a><span data-ttu-id="c0f84-122">返回值</span><span class="sxs-lookup"><span data-stu-id="c0f84-122">Return value</span></span>

```js
/** @returns {Word.Range} */
function myFunc() {

}
```

![显示 "myFunc ( # B1" 返回值的智能感知摘录的屏幕截图](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a><span data-ttu-id="c0f84-124">复杂类型</span><span class="sxs-lookup"><span data-stu-id="c0f84-124">Complex types</span></span>

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```

![显示 "var myVar;" 的复杂类型声明的 IntelliSense 的屏幕截图（示例）](../images/intellisense-vs17-complex-type.png)

## <a name="see-also"></a><span data-ttu-id="c0f84-126">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c0f84-126">See also</span></span>

- [<span data-ttu-id="c0f84-127">使用 Visual Studio 开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="c0f84-127">Develop Office Add-ins with Visual Studio</span></span>](develop-add-ins-visual-studio.md)
- [<span data-ttu-id="c0f84-128">在 Visual Studio 中调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="c0f84-128">Debug Office Add-ins in Visual Studio</span></span>](debug-office-add-ins-in-visual-studio.md)

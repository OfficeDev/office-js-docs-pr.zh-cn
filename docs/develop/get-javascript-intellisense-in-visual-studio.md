---
title: 在 Visual Studio 2017 中获取 JavaScript IntelliSense
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 78a774397069d0c6ff91cc098cad0fd9b8e5c7b9
ms.sourcegitcommit: d34aa0b282cc76ffff579da2a7945efd12fb7340
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/05/2019
ms.locfileid: "36769545"
---
# <a name="get-javascript-intellisense-in-visual-studio-2017"></a><span data-ttu-id="498ca-102">在 Visual Studio 2017 中获取 JavaScript IntelliSense</span><span class="sxs-lookup"><span data-stu-id="498ca-102">Get JavaScript IntelliSense in Visual Studio 2017</span></span>

<span data-ttu-id="498ca-p101">当使用 Visual Studio 2017 开发 Office 外接程序时，可以使用 JSDoc 来启用 IntelliSense，以获取 JavaScript 变量、对象、参数和返回值。本文概述了 JSDoc 以及如何使用它在 Visual Studio 中创建 IntellSense。有关详细信息，请参阅 [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) 和 [JavaScript 中的 JSDoc 支持](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript)。</span><span class="sxs-lookup"><span data-stu-id="498ca-p101">When you use Visual Studio 2017 to develop Office Add-ins, you can use JSDoc to enable IntelliSense for your JavaScript variables, objects, parameters, and return values. This article provides an overview of JSDoc and how you can use it to create IntellSense in Visual Studio. For more details, see [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) and [JSDoc support in JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript).</span></span> 

## <a name="officejs-type-definitions"></a><span data-ttu-id="498ca-106">Office.js 类型定义</span><span class="sxs-lookup"><span data-stu-id="498ca-106">Office.js type definitions</span></span>

<span data-ttu-id="498ca-p102">需要向 Visual Studio 提供 Office.js 中的类型定义。为此，可以执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="498ca-p102">You need to provide the definitions of the types in Office.js to Visual Studio. To do this, you can:</span></span>

- <span data-ttu-id="498ca-p103">在名为 `\Office\1\` 的解决方案的某个文件夹中，创建 Office.js 文件的本地副本。在创建外接程序项目时，Visual Studio 中的 Office 外接程序项目模板会添加此本地副本。</span><span class="sxs-lookup"><span data-stu-id="498ca-p103">Have a local copy of the Office.js files in a folder in your solution named `\Office\1\`. The Office Add-in project templates in Visual Studio add this local copy when you create an add-in project.</span></span> 
- <span data-ttu-id="498ca-p104">通过将 tsconfig.json 文件添加到外接程序解决方案中的 Web 应用程序项目的根，来使用 Office.js 的在线版本。文件应当包括以下内容。</span><span class="sxs-lookup"><span data-stu-id="498ca-p104">Use an online version of Office.js by adding a tsconfig.json file to the root of the web application project in the add-in solution. The file should include the following content.</span></span>

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

## <a name="jsdoc-syntax"></a><span data-ttu-id="498ca-113">JSDoc 语法</span><span class="sxs-lookup"><span data-stu-id="498ca-113">JSDoc syntax</span></span>

<span data-ttu-id="498ca-p105">基本技巧是在变量（或参数等）前面加上一个标识其数据类型的注释。这样，Visual Studio 中的 IntelliSense 可以推断其成员。示例如下。</span><span class="sxs-lookup"><span data-stu-id="498ca-p105">The basic technique is to precede the variable (or parameter, and so on) with a comment that identifies its data type. This allows IntelliSense in Visual Studio to infer its members. The following are examples.</span></span>

### <a name="variable"></a><span data-ttu-id="498ca-117">变量</span><span class="sxs-lookup"><span data-stu-id="498ca-117">Variable</span></span>

```js
/** @type {Excel.Range} */
var subsetRange;
```
![变量的 Intellisense](../images/intellisense-vs17-var.png)

### <a name="parameter"></a><span data-ttu-id="498ca-119">参数</span><span class="sxs-lookup"><span data-stu-id="498ca-119">Parameter</span></span>

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```
![参数的 Intellisense](../images/intellisense-vs17-param.png)

### <a name="return-value"></a><span data-ttu-id="498ca-121">返回值</span><span class="sxs-lookup"><span data-stu-id="498ca-121">Return value</span></span>

```js
/** @returns {Word.Range} */
function myFunc() {

}
```
![返回值的 IntelliSense](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a><span data-ttu-id="498ca-123">复杂类型</span><span class="sxs-lookup"><span data-stu-id="498ca-123">Complex types</span></span>

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```
![对复杂类型使用 Intellisense](../images/intellisense-vs17-complex-type.png)

## <a name="see-also"></a><span data-ttu-id="498ca-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="498ca-125">See also</span></span>

- [<span data-ttu-id="498ca-126">在 Visual Studio 中创建和调试加载项</span><span class="sxs-lookup"><span data-stu-id="498ca-126">Create and debug add-ins in Visual Studio</span></span>](create-and-debug-office-add-ins-in-visual-studio.md)

---
title: 在 Visual Studio 2019 中获取 JavaScript IntelliSense
description: 了解如何使用 JSDoc 为 JavaScript IntelliSense、对象、参数和返回值创建属性。
ms.date: 10/11/2019
ms.localizationpriority: medium
ms.openlocfilehash: 142f4cf81f83cdc518fb94f65a8e6f67f41a4783
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743490"
---
# <a name="get-javascript-intellisense-in-visual-studio-2019"></a>在 Visual Studio 2019 中获取 JavaScript IntelliSense

当使用 Visual Studio 2019 开发 Office 外接程序时，可以使用 JSDoc 来启用 IntelliSense，以获取 JavaScript 变量、对象、参数和返回值。本文概述了 JSDoc 以及如何使用它在 Visual Studio 中创建 IntellSense。有关详细信息，请参阅 [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) 和 [JavaScript 中的 JSDoc 支持](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript)。 

## <a name="officejs-type-definitions"></a>Office.js 类型定义

需要向 Visual Studio 提供 Office.js 中的类型定义。为此，可以执行下列操作：

- 在名为 `\Office\1\` 的解决方案的某个文件夹中，创建 Office.js 文件的本地副本。在创建外接程序项目时，Visual Studio 中的 Office 外接程序项目模板会添加此本地副本。 
- 通过将 tsconfig.json 文件添加到外接程序解决方案中的 Web 应用程序项目的根，来使用 Office.js 的在线版本。文件应当包括以下内容。

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

## <a name="jsdoc-syntax"></a>JSDoc 语法

基本技巧是在变量（或参数等）前面加上一个标识其数据类型的注释。这样，Visual Studio 中的 IntelliSense 可以推断其成员。示例如下。

### <a name="variable"></a>变量

```js
/** @type {Excel.Range} */
var subsetRange;
```

![Screenshot showing excerpt of IntelliSense for 'subsetRange' variable.](../images/intellisense-vs17-var.png)

### <a name="parameter"></a>参数

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```

![显示 JavaScript 示例 IntelliSense"paragraphs"参数中"paras"参数 (摘录的屏幕截图) 。](../images/intellisense-vs17-param.png)

### <a name="return-value"></a>返回值

```js
/** @returns {Word.Range} */
function myFunc() {

}
```

![屏幕截图显示"myFunc #IntelliSense () "返回值的值的摘录。](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a>复杂类型

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```

![Screenshot showing IntelliSense for complex type declaration of 'var myVar;' for example.](../images/intellisense-vs17-complex-type.png)

## <a name="see-also"></a>另请参阅

- [使用 Visual Studio 开发 Office 加载项](develop-add-ins-visual-studio.md)
- [在 Visual Studio 中调试 Office 加载项](debug-office-add-ins-in-visual-studio.md)

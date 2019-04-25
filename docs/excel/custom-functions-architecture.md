---
ms.date: 03/20/2019
description: 了解 Excel 自定义函数的运行时。
title: 自定义函数体系结构（预览版）
localization_priority: Priority
ms.openlocfilehash: b3f3d6c5eda51639a734c6d0f162c596f0c1e41b
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448600"
---
# <a name="custom-functions-architecture"></a>自定义函数体系结构

 自定义函数具有自己独特的运行时，可以优先执行计算。 本文将介绍自定义函数运行时与基于浏览器的 JavaScript 引擎之间的差异，该引擎支持加载项的其他绝大部分。

## <a name="custom-functions-runtime"></a>自定义函数运行时

Office Web 加载项可以作为任务窗格或内容窗格与用户进行交互，并且可以包括命令和自定义函数。 所有这些部分都在浏览器引擎运行时中运行，自定义函数除外。 自定义函数在单独的自定义函数运行时中运行，以优化计算速度。

请注意，如果你使用 [Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来生成项目，则自定义函数运行时将通过 functions.html 文件中引用的 custom-functions.js 脚本文件加载。 functions.html 仅用于加载运行时，且不应用作加载项的任务窗格。

下表突出显示了自定义函数运行时与浏览器引擎运行时之间的差异：

| 自定义函数运行时  | 浏览器引擎运行时    |
|------------------------------------------------------------------ |-------------------------------------------------------------------------------------------------------------- |
| 支持从单元格中返回值    | 支持 Office.js API 和 UI 元素   |
| 没有 `localStorage` 对象，改用 `AsyncStorage`  | 具有 `localStorage` 对象，可以选择使用 `AsyncStorage` 对象   |
| 不支持与 DOM 交互，或者加载依赖于 DOM 的库，如 jQuery。    | 支持与 DOM 交互，和加载依赖于 DOM 的库。 |


## <a name="browser-engine-runtime"></a>浏览器引擎运行时

任务窗格、内容加载项和命令在浏览器引擎运行时中运行。

浏览器引擎运行时支持 Office.js API。 请记住，任何 Excel API（例如允许你操作 Excel 表的 API）都可以在浏览器引擎运行时上运行，但无法从自定义函数运行时直接访问。

## <a name="communicate-between-runtimes"></a>运行时之间的通信

你的自定义函数代码无法直接与 Web 加载项的其他部分（例如任务窗格）中的代码进行交互，因为它们位于不同的运行时。 但在某些方案中，可能需要共享数据，例如传递令牌。

`AsyncStorage` 可用于存储自定义函数的数据并从任务窗格代码中获取数据。 有关存储和共享数据的详细信息，请参阅[保存和共享状态](custom-functions-overview.md#saving-and-sharing-state)。

可以使用这一专用于模式和做法的 [Github 存储库](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)中的 `AsyncStorage` 查看代码示例。
有关 `AsyncStorage` 的更多常规信息，请参阅[自定义函数运行时](./custom-functions-runtime.md)。

`AsyncStorage` 也可用于身份验证。 有关详细信息，请参阅[自定义函数身份验证](custom-functions-authentication.md)。

## <a name="see-also"></a>另请参阅

* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [自定义函数元数据](custom-functions-json.md)
* [自定义函数更改日志](custom-functions-changelog.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)

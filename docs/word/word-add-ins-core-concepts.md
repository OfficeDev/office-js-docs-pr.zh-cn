---
title: Word JavaScript API 基本编程概念
description: 使用 Word JavaScript API 生成适用于 Word 的加载项。
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: 1e7a90d4be378ed9b2c1f30ebebd4a0beec45a11
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293091"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a>Word JavaScript API 基本编程概念

本文介绍使用 [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) 生成适用于 Word 2016 或更高版本的加载项的基本概念。

## <a name="referencing-officejs"></a>引用 Office.js

可以从以下位置引用 Office.js：

- `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - 将此资源用于生产外接程序。

- `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - 通过此资源试用预览版功能。

## <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API 要求集

要求集是指各组已命名的 API 成员。 Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。 有关 Word JavaScript API 要求集的详细信息，请参阅 [Word JavaScript API 要求集](../reference/requirement-sets/word-api-requirement-sets.md)。

## <a name="running-word-add-ins"></a>运行 Word 加载项

若要运行加载项，请使用 `Office.initialize` 事件处理程序。 若要详细了解如何初始化加载项，请参阅[了解 API](../develop/understanding-the-javascript-api-for-office.md)。

面向 Word 2016 或更高版本的加载项可以使用特定于 Word 的 API。 它们将 Word 交互逻辑作为函数传递到 `Word.run()` 方法中。 请参阅[使用特定于应用程序的 API 模型](../develop/application-specific-api-model.md)，了解如何与此编程模型中的 Word 文档进行交互。

以下示例显示如何使用 `Word.run()` 方法初始化和运行 Word 加载项。

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 概述](../reference/overview/word-add-ins-reference-overview.md)
- [生成首个 Word 加载项](../quickstarts/word-quickstart.md)
- [Word 加载项教程](../tutorials/word-tutorial.md)
- [Word JavaScript API 参考](/javascript/api/word)

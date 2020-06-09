---
title: Word JavaScript API 基本编程概念
description: 使用 Word JavaScript API 生成适用于 Word 的加载项。
ms.date: 07/05/2019
localization_priority: Priority
ms.openlocfilehash: 697f3068a039caa8ae60ed449bacb05f3999a1ee
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608563"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a>Word JavaScript API 基本编程概念

本文介绍使用 [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) 生成适用于 Word 2016 或更高版本的加载项的基本概念。

## <a name="referencing-officejs"></a>引用 Office.js

可以从以下位置引用 Office.js：

- `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - 将此资源用于生产外接程序。

- `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - 通过此资源试用预览版功能。

## <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API 要求集

要求集是指各组已命名的 API 成员。 Office 外接程序使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持外接程序所需的 API。 有关 Word JavaScript API 要求集的详细信息，请参阅 [Word JavaScript API 要求集](../reference/requirement-sets/word-api-requirement-sets.md)。

## <a name="running-word-add-ins"></a>运行 Word 加载项

若要运行加载项，请使用 `Office.initialize` 事件处理程序。 若要详细了解如何初始化加载项，请参阅[了解 API](../develop/understanding-the-javascript-api-for-office.md)。

面向 Word 2016 或更高版本的加载项通过向 `Word.run()` 方法传递一个函数来运行。 传递到 `run` 方法的函数必须具有上下文参数。 此[上下文对象](/javascript/api/word/word.requestcontext)不同于从 Office 对象获取的上下文对象，但它同样可以用于与 Word 运行时环境交互。 此上下文对象可提供对 Word JavaScript API 对象模型的访问。 以下示例显示如何使用 `Word.run()` 方法初始化和运行 Word 加载项。

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

### <a name="asynchronous-nature-of-word-apis"></a>Word API 的异步特性

Word JavaScript API 是由 Office.js 加载的。 Word JavaScript API 改变了你与文档和段落等对象交互的方式。 Word JavaScript API 不提供用于检索和更新每个对象的单个异步 API，而是提供与在 Word 中运行的实时对象对应的“代理”JavaScript 对象。 通过同步读取和写入这些代理对象的属性，并调用对这些对象执行操作的同步方法，可与这些对象进行交互。 与代理对象的这些交互不会立即在运行的脚本中实现。 `context.sync` 方法通过执行排队的指令以及检索用于脚本的已加载 Word 对象的属性，在运行的 JavaScript 和 Office 中的真实对象之间同步状态。

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>将 Word 文档与 Word JavaScript API 代理对象进行同步

Word JavaScript API 对象模型与 Word 中的对象松散耦合。Word JavaScript API 对象是 Word 文档中对象的代理。在文档状态完成同步前，对代理对象执行的操作不会在 Word 中实现。反过来说，在文档状态完成同步前，Word 文档的状态也不会在代理对象中实现。若要同步文档状态，请运行 `context.sync()` 方法。下面的示例创建了代理正文对象以及用于在代理正文对象上加载文本属性的已排入队列命令，并使用 `context.sync()` 方法将 Word 文档正文与正文代理对象同步。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    body.load("text");

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a>执行一批命令

Word 代理对象具有访问和更新对象模型的方法。 这些方法按其在批处理中的排队顺序依次运行。 调用 `context.sync()` 时，批处理中已排队的所有命令都会运行。

以下示例将说明命令队列的工作原理。 调用 `context.sync()` 时，用于加载正文文本的命令会在 Word 中运行。 然后，用于在正文中插入文本的命令会在 Word 中执行。 结果将返回到 body 代理对象。 Word JavaScript API 中 `body.text` 属性的值为在将文本插入 Word 文档<u>之前</u> Word 文档正文的值。

```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    body.load("text");

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 概述](../reference/overview/word-add-ins-reference-overview.md)
- [生成首个 Word 加载项](../quickstarts/word-quickstart.md)
- [Word 加载项教程](../tutorials/word-tutorial.md)
- [Word JavaScript API 参考](/javascript/api/word)
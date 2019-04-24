---
title: Word JavaScript API 概述
description: ''
ms.date: 03/19/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 19e3b7732fb5372228ea1458c57df5e79b08078a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450091"
---
# <a name="word-javascript-api-overview"></a>Word JavaScript API 概述

Word 提供了一组丰富的 API，你可以使用它们创建与文档内容和元数据进行交互的外接程序。使用这些 API 可以为用户带来与 Word 融为一体并扩展 Word 的精彩体验。你可以导入和导出内容、组合来自不同数据源的新文档，并能与文档工作流进行集成，从而创建自定义文档解决方案。

你可以使用以下两个 JavaScript API 与 Word 文档中的对象和元数据进行交互：

- Word JavaScript API - 在 Office 2016 中引入。
- [适用于 Office 的 JavaScript API](../javascript-api-for-office.md) (Office.js) - 在 Office 2013 中引入。

## <a name="word-javascript-api"></a>Word JavaScript API

Word JavaScript API 通过 Office.js 进行加载，改变了你与文档和段落等对象的交互方式。Word JavaScript API 不提供各个用于检索和更新每个对象的异步 API，而是提供与 Word 中运行的真实对象对应的“代理”JavaScript 对象。你可以通过同步读取和写入这些代理对象的属性，并调用对其执行操作的同步方法，从而与这些代理对象进行交互。与代理对象的这些交互不会立即在运行的脚本中实现。**context.sync** 方法通过执行已排入队列的指令并检索可供在脚本中使用的已加载 Word 对象的属性，在运行的 JavaScript 和 Office 真实对象之间同步状态。

## <a name="javascript-api-for-office"></a>适用于 Office 的 JavaScript API

可以从以下位置引用 Office.js：

* https://appsforoffice.microsoft.com/lib/1/hosted/office.js - 将此资源用于生产外接程序。
* https://appsforoffice.microsoft.com/lib/beta/hosted/office.js - 在试用预览功能时使用此资源。

如果你使用的是 [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs)，则可以下载 [Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs.aspx)，从而获取包含 Office.js 的项目模板。你还可以使用 [nuget 获取 Office.js](https://www.nuget.org/packages/Microsoft.Office.js/)。

如果你使用的是 TypeScript 并且拥有 npm，则可以在命令行接口中键入以下命令，从而获取 TypeScript 定义：`typings install office-js --ambient`。

## <a name="running-word-add-ins"></a>运行 Word 外接程序

若要运行外接程序，请使用 Office.initialize 事件处理程序。若要详细了解如何初始化外接程序，请参阅[了解 API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。

面向 Word 2016 或更高版本的外接程序通过向 **Word.run()** 方法传递函数来执行。 传递到 **run** 方法的函数必须具有上下文参数。 此[上下文对象](/javascript/api/word/word.requestcontext)不同于从 Office 对象获取的上下文对象，但它同样可以用于与 Word 运行时环境交互。 此上下文对象可提供对 Word JavaScript API 对象模型的访问。 以下示例显示如何使用 **Word.run()** 方法初始化和运行 Word。

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

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>将 Word 文档与 Word JavaScript API 代理对象进行同步

Word JavaScript API 对象模型与 Word 中的对象松散耦合。Word JavaScript API 对象是 Word 文档中对象的代理。在文档状态完成同步前，对代理对象执行的操作不会在 Word 中实现。反过来说，在文档状态完成同步前，Word 文档的状态也不会在代理对象中实现。若要同步文档状态，请运行 **context.sync()** 方法。下面的示例创建了代理正文对象以及用于在代理正文对象上加载文本属性的已排入队列命令，并使用 **context.sync()** 方法将 Word 文档正文与正文代理对象同步。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a>执行一批命令

Word 代理对象具有用于访问和更新对象模型的方法。这些方法按其在批处理中的排队顺序依次执行。调用 context.sync() 后，批处理中已排入队列的所有命令都会得到执行。

下面的示例展示了命令队列的工作原理。调用 **context.sync()** 时，用于加载正文文本的命令会在 Word 中执行。然后，用于在正文中插入文本的命令会在 Word 中执行。接下来，结果会返回到正文代理对象。Word JavaScript API 中 **body.text** 属性的值为在将文本插入 Word 文档<u>之前</u> Word 文档正文的值。


```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    context.load(body, 'text');

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

## <a name="word-javascript-api-open-specifications"></a>Word JavaScript API 开放性规范

在我们设计和开发新的 API 以用于创建 Word 外接程序时，我们会公开它们，以便你可以在我们的[开放性 API 规范](../openspec.md)页面上提供反馈。了解即将推出的面向 Word JavaScript API 的新功能，并提供你对我们的设计规范的宝贵意见。

## <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API 要求集

要求集是指各组已命名的 API 成员。 Office 外接程序使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持外接程序所需的 API。 有关 Word JavaScript API 要求集的详细信息，请参阅 [Word JavaScript API 要求集](../requirement-sets/word-api-requirement-sets.md)文章。

## <a name="word-javascript-api-reference"></a>Word JavaScript API 参考

有关 Word JavaScript API 的详细信息，请参阅 [Word JavaScript API 参考文档](/javascript/api/word)。

## <a name="see-also"></a>另请参阅

* [Word 外接程序概述](/office/dev/add-ins/word/word-add-ins-programming-overview)
* [Office 外接程序平台概述](/office/dev/add-ins/overview/office-add-ins)
* [GitHub Word 上的外接程序示例](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)

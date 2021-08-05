---
title: 初始化 Office 加载项
description: 了解如何初始化Office外接程序。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 0cddc4eaa99c9f1536be91d6fe2971c43344a149
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/05/2021
ms.locfileid: "53774292"
---
# <a name="initialize-your-office-add-in"></a>初始化 Office 加载项

Office 加载项通常使用启动逻辑执行以下操作：

- 检查用户版本是否支持Office调用的所有 Office API。

- 确保存在某些项目，例如具有特定名称的工作表。

- 提示用户选择文档中的一些Excel，然后插入用这些选定值初始化的图表。

- 建立绑定。

- 使用 Office 对话框 API 提示用户输入默认外接程序设置值。

但是，Office加载库之前，Office无法成功调用任何 JavaScript API。 本文介绍了代码可以确保已加载库的两种方法。

- 使用 进行初始化 `Office.onReady()` 。
- 使用 进行初始化 `Office.initialize` 。

> [!TIP]
> 建议使用 `Office.onReady()` 取代 `Office.initialize`。 尽管 `Office.initialize` 仍受支持， `Office.onReady()` 但提供了更大的灵活性。 只能向 分配一个处理程序，并且该处理程序仅由 Office `Office.initialize` 调用一次。 可以在代码中 `Office.onReady()` 的不同位置调用 ，并使用不同的回调。
> 
> 有关这两种方法之间的差别信息，请参阅 [Office.initialize 和 Office.onReady() 之间的主要差别](#major-differences-between-officeinitialize-and-officeonready)。

有关初始化加载项时的事件顺序的更多详细信息，请参阅 [加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。

## <a name="initialize-with-officeonready"></a>使用 Office.onReady() 进行初始化

`Office.onReady()` 是一种异步方法 [，在检查](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) 是否加载 Promise 库时Office.js Promise 对象。 加载库时，它会将 Promise 解析为一个对象，该对象使用枚举值 (、 等 ) 指定 Office 客户端应用程序，以及具有枚举值 `Office.HostType` `Excel` `Word` `Office.PlatformType` `PC` `Mac` (、、 `OfficeOnline` 等 ) 的平台。 如果在调用 `Office.onReady()` 时已加载库，则 Promise 将立即解析。

调用 `Office.onReady()` 的一种方法是向其传递一个回调方法。 下面是一个示例。

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

或者，可以将 `then()` 方法链接到 `Office.onReady()` 的调用，而不是传递回调。 例如，以下代码检查用户的 Excel 版本是否支持加载项可能调用的所有 API。

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

下面是在 TypeScript 中使用 `async` 和 `await` 关键字的相同示例。

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

如果使用的是其他 JavaScript 框架，其中包括它们自己的初始化处理程序或测试，那么它们 *通常* 应放置在 `Office.onReady()` 的响应内。 例如，会对 [JQuery 的](https://jquery.com) `$(document).ready()` 函数进行以下引用：

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

但是，此做法存在例外情况。 例如，假设你想要在浏览器 (而不是在 Office 应用程序) 中旁加载外接程序，以便使用浏览器工具调试 UI。 由于 Office.js 将不会在浏览器中加载，所以，`onReady` 将不会运行，且如果在 Office `$(document).ready` 内调用它，则 `onReady` 将不会运行。 

另一个例外情况是，您希望在加载外接程序时在任务窗格中显示进度指示器。 在此方案中，代码应调用 jQuery 并使用 `ready` 其回调来呈现进度指示器。 然后，Office `onReady` 的回调可将进度指示器替换为最终 UI。 

## <a name="initialize-with-officeinitialize"></a>使用 Office.initialize 进行初始化

当 Office.js 库加载并准备好用于用户交互时将触发初始化事件。 可将处理程序分配到实现初始化逻辑的 `Office.initialize`。 以下是检查用户的 Excel 版本是否支持加载项可能调用的所有 API 的示例。

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

如果你使用的是包含它们自己的初始化处理程序或测试的其他 JavaScript 框架，它们通常应放置在事件 (前面使用 `Office.initialize` **Office.onReady ()** 初始化部分中介绍的异常也适用于) 。 例如，会对 [JQuery 的](https://jquery.com) `$(document).ready()` 函数进行以下引用：

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

对于任务窗格和内容加载项，`Office.initialize` 提供了其他 _reason_ 参数。 此参数指定如何将加载项添加到当前文档。 可以使用此参数针对首次插入加载项时和加载项已存在于文档中时实施不同的逻辑。

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

有关详细信息，请参阅 [Office.initialize 事件](/javascript/api/office)和 [InitializationReason 枚举](/javascript/api/office/office.initializationreason)。

## <a name="major-differences-between-officeinitialize-and-officeonready"></a>Office.initialize 和 Office.onReady 之间的主要差别

- 可以仅将一个处理程序分配到 `Office.initialize` 并仅由 Office 基础结构调用一次，但可以在代码中的不同位置调用 `Office.onReady()` 并使用不同的回调。 例如，只要自定义脚本使用运行初始化逻辑的回调进行加载，代码就可以调用 `Office.onReady()`。代码还可以在任务窗格中设置一个按钮，其脚本会使用不同的回调调用 `Office.onReady()`。 如果是这样，则会在单击该按钮后运行第二个回调。

- `Office.initialize` 事件将在 Office.js 初始化其本身的内部过程的末尾处触发。 并且它会在内部过程结束后 *立即* 触发。 如果将处理程序分配到事件所使用的代码在事件触发后执行的时间过长，则处理程序将不会运行。 例如，如果使用的是 WebPack 任务管理器，则在加载 Office.js 后但在加载自定义 JavaScript 前，它会配置加载项的主页以加载填充代码文件。 在脚本加载和分配处理程序时，初始化事件已经发生。 但调用 `Office.onReady()` 永远不会“太迟”。 如果初始化事件已经发生，则回调将立即运行。

> [!NOTE]
> 即使没有启动逻辑，也应在加载项 JavaScript 加载时调用 `Office.onReady()` 或将空函数分配到 `Office.initialize`。 某些Office应用程序组合和平台组合不会加载任务窗格，直到其中一个情况发生。 以下示例显示了这两种方法。
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)
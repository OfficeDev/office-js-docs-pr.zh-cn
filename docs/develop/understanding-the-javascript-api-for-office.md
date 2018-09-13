---
title: 了解适用于 Office 的 JavaScript API
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 3bda3975634e1d9a84684c66fcf35bd7da66f701
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945458"
---
# <a name="understanding-the-javascript-api-for-office"></a>了解适用于 Office 的 JavaScript API

本文提供了有关适用于 Office 的 JavaScript API 的信息以及使用方法。有关参考信息，请参阅 [适用于 Office 的 JavaScript API](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)。有关将 Visual Studio 项目文件更新到适用于 Office 的 JavaScript API 的最新当前版本的信息，请参阅 [更新适用于 Office 的 JavaScript API 版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。

> [!NOTE]
> 如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>在加载项中引用适用于 Office 的 JavaScript API 库

[适用于 Office 的 JavaScript](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) 库包含 Office.js 文件和关联的特定于主机应用程序的 .js 文件，例如 Excel-15.js 和 Outlook-15.js。引用该 API 最简单的方法是通过添加以下 `<script>` 到你的页面的 `<head>` 标记来使用我们的 CDN：  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

这将在加载项首次加载时下载并缓存适用于 Office 的 JavaScript API 文件，以确保对特定版本使用 Office.js 及其关联文件的最新实现。

有关 Office.js CDN 的更多详细信息（包括版本控制和向后兼容性），请参阅[从内容分发网络 (CDN) 引用适用于 Office 的 JavaScript API 库](referencing-the-javascript-api-for-office-library-from-its-cdn.md)。

## <a name="initializing-your-add-in"></a>初始化加载项

**适用于：** 所有加载项类型

Office 加载项通常有启动逻辑，以执行以下事项：

- 检查用户的 Office 版本是否支持您的代码调用的所有 Office Api。

- 确保存在某些工件，如具有特定名称的工作表。

- 提示用户选择 Excel 中的一些单元格，然后插入使用这些选定值初始化的图表。

- 建立绑定。

- 使用 Office 对话框 API 提示用户输入默认加载项设置值。

但是，在完全加载完库之前，您启动代码不得调用任何 Office.js Api。 有两种方法让您的代码可以确保加载库。 这将在以下各节介绍。 我们建议您使用名为 `Office.onReady()` 的较新、 更灵活的技术。 仍然支持分配处理程序 `Office.initialize` 的旧技术。 请参阅 [Office.initialize 和 Office.onReady() 的主要区别](#major-differences-between-office-initialize-and-office-onready)。

若要详细了解加载项初始化时的事件发生顺序，请参阅[加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。

### <a name="initialize-with-officeonready"></a>使用 Office.onReady() 初始化

`Office.onReady()` 是返回承诺对象，同时检查 Office.js 库是否完全加载的异步方法。 只有在加载库后，它才会将承诺解析为对象，这将使用`Office.HostType` 枚举值 (`Excel`， `Word`等) 和与平台 `Office.PlatformType` 枚举值 (`PC`， `Mac`， `OfficeOnline`，等）指定 Office 主机应用程序。 如果在调用 `Office.onReady()` 时已加载库，则承诺立即解析。

调用的一种方法 `Office.onReady()` 是，将其传递给回调方法。 下面是一个示例：

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

或者，您可以将 `then()` 方法与 `Office.onReady()` 的调用链接而不是传递回调。 例如，下面的代码将检查用户的 Excel 版本是否支持加载项可能调用的所有 Api。

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

以下是在 TypeScript 中使用 `async` 和 `await` 关键字的相同示例：

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

如果您正使用包括自有初始化处理程序或测试的其他 JavaScript 框架，则这些框架*通常应*位于 `Office.onReady()` 的响应中。 例如，会对 [JQuery](https://jquery.com) `$(document).ready()` 函数进行以下引用：

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

但是，这一做法存在一些例外。 例如，假设您想要在浏览器中打开您的加载项（而不是 侧加载到 Office 主机），从而使用浏览器工具调试您的 UI。 由于 Office.js 无法在浏览器中加载，`onReady` 将无法运行，同时如果在 Office `onReady` 内调用，`$(document).ready` 将无法运行。 另一个异常：加载加载项期间，您希望在任务窗格中显示进度指示器。 在此方案中，您的代码应调用 jQuery `ready`，并使用它的回调以呈现进度指示器。 然后，Office `onReady`的回调可以替换最终用户界面的进度指示器。 

### <a name="initialize-with-officeinitialize"></a>使用 Office.initialize 初始化

当 Office.js 库完全加载并可供用户交互时，初始化事件触发。 您可以分配一个处理程序给实施初始化逻辑的 `Office.initialize`。 以下是检查用户的 Excel 版本是否支持所有可能调用加载项的 Api 示例。

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

如果您正使用包括自有初始化处理程序或测试的其他 JavaScript 框架，则这些框架*通常*应位于 `Office.initialize` 事件内。 （但是，更早版本 **与 Office.onReady() 初始化** 一节描述的异常也适用于这种情况。）例如， [JQuery](https://jquery.com) `$(document).ready()`函数会被引用为：

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

对于任务窗格和内容加载项，`Office.initialize` 提供其他_原因_参数。 此参数指定如何添加加载项到当前文档。 你可以使用此参数针对首次插入加载项时和加载项已存在于文档中时实施不同的逻辑。

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

有关详细信息，请参阅 [Office.initialize 事件](https://docs.microsoft.com/javascript/api/office?view=office-js)和 [InitializationReason 枚举](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js)。

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Office.initialize 和 Office.onReady 的主要区别

- 您仅可分配一个处理程序到 `Office.initialize`，同时它由由 Office 基础架构仅调用一次；但是，您可以在代码中的不同位置调用 `Office.onReady()` 并可使用不同的回调。 例如，一旦使用运行初始化逻辑的调用加载您的自定义脚本，您的代码即可调用 `Office.onReady()`；同时，您的代码还可在任务窗格中有一个按钮，其脚本以不同的回调来调用 `Office.onReady()`。 如果是这样，单击按钮时将运行第二个回调。

- 事件在 Office.js 初始化本身的内部过程末尾触发。`Office.initialize` 这在内部过程结束后*立即*触发。 如果事件触发后指定处理程序给事件的代码执行时间过长，则不运行您的处理程序。 例如，如果您使用 WebPack 任务管理器，它可能在加载 Office.js 后，但在加载您的自定义 JavaScript 之前配置加载项主页以加载 polyfill 文件。 脚本加载并分配该处理程序时，初始化事件已经发生。 但是，调用 `Office.onReady()` 不会"过晚"。 如果初始化事件已经发生，则回调立即运行。

> [!NOTE]
> 即使未启动逻辑，当加载加载项 JavaScript 时，调用 `Office.onReady()` 或分配到一个空函数给 `Office.initialize` 是一个好的做法，因为在发生下列任一情况之前，某些 Office 主机和平台组合不会加载任务窗格。 以下各行显示可以完成这个的两种方式：
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Office JavaScript API 对象模型

初始化后，加载项可以与主机（例如 Excel、Outlook）交互。 [Office JavaScript API 对象模型](office-javascript-api-object-model.md)页面有关于特定使用模式的更多详细信息。 [共享 API](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) 和特定的主机都有详细的参考文档。

## <a name="api-support-matrix"></a>API 支持矩阵

下表总结了各种类型的加载项（内容、任务窗格和 Outlook）支持的 API 和功能，以及使用[适用于 Office 的 JavaScript API v1.1 支持的 1.1 加载项清单架构和功能](update-your-javascript-api-for-office-and-manifest-schema-version.md)指定加载项支持的 Office 主机应用时，可以托管它们的 Office 应用。


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**主机名**|数据库|工作簿|邮箱|演示文稿|文档|项目|
||**支持的****主机应用程序**|Access Web App|Excel、<br/>Excel 在线|Outlook、<br/>Outlook Web App、<br/>适用于设备的 OWA|PowerPoint、<br/>PowerPoint 联机|Word|项目|
|**支持的外接程序类型**|内容|是|是||是|||
||任务窗格||是||是|是|是|
||Outlook|||是||||
|**支持的 API 功能**|读/写文本||是||是|是|是<br/>（只读）|
||读/写矩阵||是|||是||
||读/写表||是|||是||
||读/写 HTML|||||是||
||读/写<br/>Office Open XML|||||是||
||读取任务、资源、视图和字段属性||||||是|
||选择已更改事件||是|||是||
||获取整个文档||||是|是||
||绑定和绑定事件|是<br/>（仅限完全和部分表格绑定）|是|||是||
||读/写自定义 XML 部分|||||是||
||暂留加载项状态数据（设置）|是<br/>（每主机加载项）|是<br/>（每文档）|是<br/>（每邮箱）|是<br/>（每文档）|是<br/>（每文档）||
||设置更改事件|是|是||是|是||
||获取活动视图模式<br/>和视图更改事件||||是|||
||转到文档中<br/>的相应位置||是||是|是||
||使用规则和 RegEx<br/>执行上下文式激活|||是||||
||读取项目属性|||是||||
||读取用户配置文件|||是||||
||获取附件|||是||||
||获取用户标识令牌|||是||||
||调用 Exchange Web 服务|||是||||

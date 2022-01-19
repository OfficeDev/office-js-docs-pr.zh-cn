---
title: 支持 Internet Explorer 11
description: 了解如何在外接程序Internet Explorer 11 和 ES5 Javascript。
ms.date: 10/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: 755bcde8748b3cc0ce2f5de92a6ba5f04f6d263c
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074257"
---
# <a name="support-internet-explorer-11"></a>支持 Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer外接程序Office中使用的内容**
>
> Microsoft 将终止对Internet Explorer的支持，但这不会显著影响Office外接程序。平台和 Office 版本（包括 Office 2019 之间的一次购买版本）的一些组合将继续使用 Internet Explorer 11 随附的 Webview 控件来托管外接程序，如[Office](../concepts/browsers-used-by-office-web-add-ins.md)外接程序使用的浏览器所说明。此外，提交到[AppSource](/office/dev/store/submit-to-appsource-via-partner-center)的加载项仍然需要支持这些Internet Explorer，因此，对于加载项，这些组合也仍是必需的。 有两 *个变化* ：
>
> - Office web 版中不再打开Internet Explorer。 因此，AppSource 不再使用浏览器Office web 版Internet Explorer加载项。 但 AppSource 仍测试使用 Office *版本的平台* 和桌面Internet Explorer。
> - Script Lab[工具](../overview/explore-with-script-lab.md)不再支持Internet Explorer。

Office外接程序是 Web 应用程序，当在 IFrame 上运行时，这些应用程序Office web 版。 Office在 Mac 上的 Office 或 Windows Office 中运行时，外接程序使用嵌入式浏览器控件显示。 嵌入式浏览器控件由操作系统或用户计算机上安装的浏览器提供。

如果计划通过 AppSource 销售加载项或计划支持较旧版本的 Windows 和 Office，加载项必须在基于 Internet Explorer 11 (IE11) 的可嵌入浏览器控件中运行。 有关使用基于 IE11 Windows和Office的浏览器控件的信息，请参阅 Office [Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)使用的浏览器。

> [!IMPORTANT]
> Internet Explorer 11 不支持某些 HTML5 功能，如媒体、录制和位置。 如果外接程序必须支持 Internet Explorer 11，则必须设计外接程序以避免这些不受支持的功能，或者外接程序必须检测 Internet Explorer 何时使用，并提供不使用不受支持功能的备用体验。 有关详细信息，请参阅在运行时确定 [加载项](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)是否正在Internet Explorer。

## <a name="support-for-recent-versions-of-javascript"></a>支持 JavaScript 的最新版本

Internet Explorer 11 不支持低于 ES5 的 JavaScript 版本。 如果要使用 ECMAScript 2015 或更高版本或 TypeScript 的语法和功能，有两个选项，如本文所述。 还可以结合这两种技术。

### <a name="use-a-transpiler"></a>使用转译器

可以使用 TypeScript 或新式 JavaScript 编写代码，然后在生成时将代码转换为 ES5 JavaScript。 生成的 ES5 文件是上传到加载项 Web 应用程序的文件。

有两种常用转译器。 两者都可以使用 TypeScript 或 ES5 后 JavaScript 的源文件。 它们还使用 React.jsx ( .tsx) 。

- [一些](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

有关在加载项项目中安装和配置转译器的信息，请参阅任一文档。 建议您使用任务运行程序（如 [Grunt](https://gruntjs.com/) 或 [WebPack）](https://webpack.js.org/) 来自动进行转换。 有关使用 tsc 的示例外接程序，请参阅 Office Microsoft 外接程序[Graph React。](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React) 有关使用分贝的示例，请参阅脱机[存储外接程序。](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin)

> [!NOTE]
> 如果使用的不是Visual Studio (，Visual Studio Code) ，则 tsc 可能最易于使用。 可以使用 nuget 程序包安装对它的支持。 有关详细信息，请参阅[JavaScript and TypeScript in Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019)。 若要对任务Visual Studio，请创建生成脚本或使用 Visual Studio 中的任务运行程序资源管理器以及[WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner)任务运行程序或[NPM 任务运行程序等工具](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner)。

### <a name="use-a-polyfill"></a>使用填充

[填充是](https://en.wikipedia.org/wiki/Polyfill_(programming))早期版本的 JavaScript，与 JavaScript 的较新版本重复功能。 填充适用于不支持更高版本 JavaScript 的浏览器。 例如，字符串方法不是 ES5 版本的 JavaScript 的一部分，因此它不会在 `startsWith` 11 Internet Explorer中运行。 有一些用 ES5 编写的填充库定义了和实现 `startsWith` 一个方法。 我们建议使用 [core-js](https://github.com/zloirock/core-js) 填充库。

若要使用填充库，请像加载任何其他 JavaScript 文件或模块一样加载它。 例如，您可以使用外接程序主页 HTML 文件 (例如) 中的 标记，或者可以使用 JavaScript 文件 (例如) 中的 `<script>` `<script src="/js/core-js.js"></script>` `import` `import 'core-js';` 语句。 当 JavaScript 引擎看到类似 的方法时，它将首先查看语言中是否内置了该 `startsWith` 名称的方法。 如果存在，它将调用本机方法。 如果且仅在该方法未内置时，引擎将查找它的所有加载文件。 因此，填充版本不会在支持本机版本的浏览器中使用。

导入整个 core-js 库将导入所有 core-js 功能。 还可以仅导入加载项Office填充。 有关如何执行此操作的说明，请参阅[CommonJS API。](https://github.com/zloirock/core-js#commonjs-api) core-js 库具有所需的大部分填充。 core-js 文档的"缺少 [填充](https://github.com/zloirock/core-js#missing-polyfills) "部分详细介绍了一些例外情况。 例如，它不支持 `fetch` ，但可以使用 [fetch](https://github.com/github/fetch) 填充。

有关使用加载项示例core.js，请参阅 [Word 加载项 Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)。

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>在运行时确定加载项是否正在Internet Explorer

通过读取 [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) 属性，Internet Explorer加载项能否在外接程序中运行。 这使外接程序能够提供备用体验或正常失败。 示例如下。 请注意，Internet Explorer以"Trident"开头的字符串作为 userAgent 的值。

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade to 
    //      either one-time purchase Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> 通常，读取属性不是一种 `userAgent` 好的做法。 请务必熟悉使用用户代理的浏览器检测一[](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent)文，包括阅读的建议和备选方法 `userAgent` 。 特别是，如果您在上述子句中采用选项 1，请考虑使用 `else` 功能检测，而不是测试用户代理。
>
> 自 2021 年 9 月 30 日起，"用户代理的哪个部分包含要查找的信息 [？"](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) 部分的文本自 Internet Explorer 11 发布之前的日期开始。 它通常仍然准确，并且本文英文部分中的表格是最新的。 同样，本文非英语版本的文本和多数情况下的表格都已过期。

## <a name="test-an-add-in-on-internet-explorer"></a>在加载项上测试Internet Explorer

请参阅 [Internet Explorer 11 测试](../testing/ie-11-testing.md)。

## <a name="additional-resources"></a>其他资源

- [ECMAScript 6 兼容性表](https://kangax.github.io/compat-table/es6/)
- [我可以使用...HTML5、CSS3 等的支持表](https://caniuse.com/)

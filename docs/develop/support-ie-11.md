---
title: 支持 Internet Explorer 11
description: 了解如何在外接程序中支持 Internet Explorer 11 和 ES5 Javascript。
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: aff6004af4ce28aea865cb34cd34e13e23fb549f
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810272"
---
# <a name="support-internet-explorer-11"></a>支持 Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer 仍在 Office 加载项中使用**
>
> 平台和 Office 版本的一些组合（包括 Office 2019 的永久版本）仍使用 Internet Explorer 11 附带的 Webview 控件来托管加载项，如 [Office 外接程序使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)中所述。我们建议 (但不需要) ，至少在 Internet Explorer Web 视图中启动加载项时，通过向外接程序的用户提供正常失败消息，继续支持这些组合。 请记住以下附加要点：
>
> - Office web 版不再在 Internet Explorer 中打开。 因此，[AppSource](/office/dev/store/submit-to-appsource-via-partner-center) 不再使用 Internet Explorer 作为浏览器在 Office web 版 中测试加载项。
> - AppSource 仍会测试使用 Internet Explorer 的平台和 Office *桌面* 版本的组合，但仅在加载项不支持 Internet Explorer 时发出警告：AppSource 不会拒绝加载项。
> - [Script Lab工具](../overview/explore-with-script-lab.md)不再支持 Internet Explorer。

Office 外接程序是在 Office web 版 上运行时显示在 IFrame 中的 Web 应用程序。 在 Windows 上的 Office 或 Mac 上的 Office 中运行时，使用嵌入式浏览器控件显示 Office 加载项。 嵌入式浏览器控件由操作系统或用户计算机上安装的浏览器提供。

如果计划支持较旧版本的 Windows 和 Office，外接程序必须在基于 Internet Explorer 11 (IE11) 的可嵌入浏览器控件中工作。 有关哪些 Windows 和 Office 组合使用基于 IE11 的浏览器控件的信息，请参阅 [Office 外接程序使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!IMPORTANT]
> Internet Explorer 11 不支持某些 HTML5 功能，例如媒体、录制和位置。 如果外接程序必须支持 Internet Explorer 11，则必须设计加载项以避免这些不受支持的功能，或者外接程序必须检测何时使用 Internet Explorer，并提供不使用不支持的功能的备用体验。 有关详细信息，请参阅 [确定在运行时加载项是否在 Internet Explorer 中运行](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)。

## <a name="support-for-recent-versions-of-javascript"></a>对最新版本的 JavaScript 的支持

Internet Explorer 11 不支持 ES5 之后的 JavaScript 版本。 如果要使用 ECMAScript 2015 或更高版本或 TypeScript 的语法和功能，可以使用本文中所述的两个选项。 还可以将这两种技术组合在一起。

### <a name="use-a-transpiler"></a>使用转码器

可以在 TypeScript 或新式 JavaScript 中编写代码，然后在生成时将其转译为 ES5 JavaScript。 生成的 ES5 文件是上传到加载项 Web 应用程序的内容。

有两种流行的转译器。 两者都可以使用 TypeScript 或 ES5 后 JavaScript 的源文件。 它们还适用于 React 文件 (.jsx 和 .tsx) 。

- [巴贝尔](https://babeljs.io/)
- [Tsc](https://www.typescriptlang.org/index.html)

有关在外接程序项目中安装和配置转码器的信息，请参阅其中任一文档。 建议使用任务运行程序（如 [Grunt](https://gruntjs.com/) 或 [WebPack](https://webpack.js.org/) ）自动执行转译。 有关使用 tsc 的示例外接程序，请参阅 [Office 外接程序 Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)。 有关使用 babel 的示例，请参阅 [脱机存储加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin)。

> [!NOTE]
> 如果使用 Visual Studio (不Visual Studio Code) ，则 tsc 可能最容易使用。 可以使用 nuget 包安装对它的支持。 有关详细信息，请参阅 [Visual Studio 2019 中的 JavaScript 和 TypeScript](/visualstudio/javascript/javascript-in-vs-2019)。 若要将 babel 与 Visual Studio 配合使用，请创建生成脚本或使用 Visual Studio 中的任务运行器资源管理器，以及 [WebPack 任务运行器](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) 或 [NPM 任务运行器](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner)等工具。

### <a name="use-a-polyfill"></a>使用 polyfill

[polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) 是早期版本的 JavaScript，它与较新版本的 JavaScript 的功能重复。 polyfill 在不支持更高 JavaScript 版本的浏览器中使用。 例如，字符串方法 `startsWith` 不是 ES5 版本的 JavaScript 的一部分，因此它不会在 Internet Explorer 11 中运行。 有一些以 ES5 编写的 polyfill 库，用于定义和实现 `startsWith` 方法。 建议使用 [core-js](https://github.com/zloirock/core-js) polyfill 库。

若要使用 polyfill 库，请像加载任何其他 JavaScript 文件或模块一样加载它。 例如，可以在外接程序的主页 HTML 文件中使用 `<script>` 标记 (例如 `<script src="/js/core-js.js"></script>`) ，也可以在 JavaScript 文件中使用 `import` 语句， (例如 `import 'core-js';`) 。 当 JavaScript 引擎看到类似 `startsWith`的方法时，它会首先查看语言中是否内置了该名称的方法。 如果有，它将调用本机方法。 如果并且仅当方法不是内置方法时，引擎才会查找它的所有已加载文件。 因此，在支持本机版本的浏览器中不使用 polyfilled 版本。

导入整个 core-js 库将导入所有 core-js 功能。 还可以仅导入 Office 外接程序所需的填充。 有关如何执行此操作的说明，请参阅 [CommonJS API](https://github.com/zloirock/core-js#commonjs-api)。 core-js 库包含所需的大部分填充。 core-js 文档的 [缺少 Polyfills](https://github.com/zloirock/core-js#missing-polyfills) 部分详细介绍了一些例外。 例如，它不支持 `fetch`，但你可以使用 [提取](https://github.com/github/fetch) polyfill。

有关使用 core.js 的示例外接程序，请参阅 [Word 外接程序 Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)。

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>在运行时确定加载项是否在 Internet Explorer 中运行

加载项可以通过读取 [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) 属性来发现它是否在 Internet Explorer 中运行。 这使外接程序能够提供备用体验或正常失败。 示例如下。 请注意，Internet Explorer 发送以“Trident”开头的字符串作为 userAgent 的值。

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade 
    //      either to perpetual Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> 读取 属性通常不是一种好的做法 `userAgent` 。 请确保熟悉 [使用用户代理进行浏览器检测](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent)一文，包括阅读 `userAgent`的建议和替代方法。 特别是，如果采用上述子句中的 `else` 选项 1，请考虑使用功能检测而不是对用户代理进行测试。
>
> 截至 2021 年 9 月 30 日， [用户代理的哪一部分包含你正在查找的信息？](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) 部分中的文本是从 Internet Explorer 11 发布之前的日期开始的。 它通常仍然准确，并且本文英文版部分中的 *表* 是最新的。 同样，在非英语版本的文章中，文本和大多数情况下的表都已过期。

## <a name="test-an-add-in-on-internet-explorer"></a>在 Internet Explorer 上测试加载项

请参阅 [Internet Explorer 11 测试](../testing/ie-11-testing.md)。

## <a name="additional-resources"></a>其他资源

- [ECMAScript 6 兼容性表](https://kangax.github.io/compat-table/es6/)
- [是否可以使用...HTML5、CSS3 等的支持表](https://caniuse.com/)

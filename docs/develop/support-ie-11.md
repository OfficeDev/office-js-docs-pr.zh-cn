---
title: 支持 Internet Explorer 11
description: 了解如何在加载项中支持 Internet Explorer 11 和 ES5 Javascript。
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1cb641f1ed1a75fcff23291d1fa566bbf6dc008b
ms.sourcegitcommit: fb3b1c6055e664d015703623661d624251ceb6b7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/17/2022
ms.locfileid: "66136423"
---
# <a name="support-internet-explorer-11"></a>支持 Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer 仍在Office加载项中使用**
>
> 一些平台和Office版本的组合，包括到 2019 Office的一次性购买版本，仍然使用 Internet Explorer 11 附带的 Webview 控件来托管加载项，如[Office加载项使用的浏览器中](../concepts/browsers-used-by-office-web-add-ins.md)所述。建议 (但不需要) 继续支持这些组合（至少以最小方式）在 Internet Explorer Webview 中启动外接程序时为外接程序的用户提供正常故障消息。 请记住以下附加点：
>
> - Office web 版不再在 Internet Explorer 中打开。 因此，[AppSource](/office/dev/store/submit-to-appsource-via-partner-center) 不再使用 Internet Explorer 作为浏览器在Office web 版中测试加载项。
> - AppSource 仍在测试使用 Internet Explorer 的平台和Office *桌面* 版本的组合，但是仅当外接程序不支持 Internet Explorer 时才会发出警告;AppSource 不会拒绝该外接程序。
> - [Script Lab工具](../overview/explore-with-script-lab.md)不再支持 Internet Explorer。

Office加载项是在 Office web 版 上运行时显示在 IFrame 中的 Web 应用程序。 Office在 Mac 上Windows或Office上Office中运行时，使用嵌入式浏览器控件显示加载项。 嵌入式浏览器控件由操作系统或用户计算机上安装的浏览器提供。

如果计划支持较旧版本的Windows和Office，则加载项必须在基于 Internet Explorer 11 (IE11) 的可嵌入浏览器控件中工作。 有关Windows和Office使用基于 IE11 的浏览器控件的组合的信息，请参阅[Office加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!IMPORTANT]
> Internet Explorer 11 不支持某些 HTML5 功能，例如媒体、录制和位置。 如果外接程序必须支持 Internet Explorer 11，则必须设计外接程序以避免这些不受支持的功能，或者加载项必须检测何时使用 Internet Explorer，并提供不使用不受支持的功能的备用体验。 有关详细信息，请参阅 [在运行时确定外接程序是否在 Internet Explorer 中运行](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)。

## <a name="support-for-recent-versions-of-javascript"></a>支持最新版本的 JavaScript

Internet Explorer 11 不支持低于 ES5 的 JavaScript 版本。 如果要使用 ECMAScript 2015 或更高版本或 TypeScript 的语法和功能，可使用本文中所述的两个选项。 还可以结合这两种技术。

### <a name="use-a-transpiler"></a>使用转译器

可以在 TypeScript 或新式 JavaScript 中编写代码，然后在生成时将其转译为 ES5 JavaScript。 生成的 ES5 文件是上传到外接程序的 Web 应用程序的内容。

有两个流行的转译器。 两者都可以使用 TypeScript 或 帖子-ES5 JavaScript 的源文件。 它们还使用React文件 (.jsx 和 .tsx) 。

- [巴贝尔](https://babeljs.io/)
- [Tsc](https://www.typescriptlang.org/index.html)

有关在外接程序项目中安装和配置转译器的信息，请参阅其中任一文档。 建议使用任务运行程序（如 [Grunt](https://gruntjs.com/) 或 [WebPack）](https://webpack.js.org/) 自动执行转译。 有关使用 tsc 的示例加载项，请[参阅Office加载项 Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)。 有关使用 babel 的示例，请参阅[脱机存储加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin)。

> [!NOTE]
> 如果使用的Visual Studio (不是Visual Studio Code) ，则 tsc 可能最容易使用。 可以使用 nuget 包安装对它的支持。 有关详细信息，请参阅 [Visual Studio 2019 中的 JavaScript 和 TypeScript](/visualstudio/javascript/javascript-in-vs-2019)。 若要将 babel 与Visual Studio配合使用，请创建生成脚本，或者将Visual Studio中的任务运行程序资源管理器与 [WebPack 任务运行程序](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner)或 [NPM 任务运行程序](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner)等工具配合使用。

### <a name="use-a-polyfill"></a>使用 polyfill

[多填充](https://en.wikipedia.org/wiki/Polyfill_(programming))是较早版本的 JavaScript，它复制了最新版本的 JavaScript 中的功能。 多填充在不支持更高版本的 JavaScript 的浏览器中使用。 例如，字符串方法 `startsWith` 不是 ES5 版 JavaScript 的一部分，因此它不会在 Internet Explorer 11 中运行。 有一些以 ES5 编写的多填充库定义和实现 `startsWith` 方法。 建议使用 [core-js](https://github.com/zloirock/core-js) polyfill 库。

若要使用多填充库，请像加载任何其他 JavaScript 文件或模块一样加载它。 例如，可以在外接程序的主页 HTML 文件 (（例如`<script src="/js/core-js.js"></script>`) ）中使用`<script>`标记，也可以在 JavaScript 文件 (中使用`import`语句，例如 `import 'core-js';`) 。 当 JavaScript 引擎看到类似 `startsWith`的方法时，它将首先查看该语言中是否内置了该名称的方法。 如果存在，它将调用本机方法。 如果该方法不是内置的，并且仅当该方法不是内置的，则引擎将查找所有已加载的文件。 因此，在支持本机版本的浏览器中不使用多填充版本。

导入整个 core-js 库将导入所有 core-js 功能。 还可以仅导入Office外接程序所需的多文件。 有关如何执行此操作的说明，请参阅 [CommonJS API](https://github.com/zloirock/core-js#commonjs-api)。 core-js 库包含所需的大部分多文件。 core-js 文档的 [“缺少 Polyfills](https://github.com/zloirock/core-js#missing-polyfills) ”部分中详述了一些异常。 例如，它不支持 `fetch`，但可以使用 [提取](https://github.com/github/fetch) 多填充。

有关使用core.js的示例加载项，请 [参阅 Word 加载项 Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)。

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>在运行时确定加载项是否在 Internet Explorer 中运行

外接程序可以通过读取 [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) 属性来发现它是否在 Internet Explorer 中运行。 这使加载项能够提供备用体验或正常失败。 示例如下。 请注意，Internet Explorer 发送以“Trident”开头的字符串作为 userAgent 的值。

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
> 读取属性通常不是一个好的做法 `userAgent` 。 请确保你熟悉文章， [使用用户代理进行浏览器检测](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent)，包括阅读的建议和替代方法 `userAgent`。 特别是，如果在上述子句中 `else` 采用选项 1，请考虑使用功能检测，而不是对用户代理进行测试。
>
> 自 2021 年 9 月 30 日起， [用户代理的哪个部分包含要查找的信息](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) 部分中的文本？日期从 Internet Explorer 11 发布之前开始。 它通常仍然准确，并且本文英文版部分中的 *表* 是最新的。 同样，在大多数情况下，本文的非英语版本中的文本和表已过时。

## <a name="test-an-add-in-on-internet-explorer"></a>在 Internet Explorer 上测试加载项

请参阅 [Internet Explorer 11 测试](../testing/ie-11-testing.md)。

## <a name="additional-resources"></a>其他资源

- [ECMAScript 6 兼容性表](https://kangax.github.io/compat-table/es6/)
- [我可以使用...HTML5、CSS3 等的支持表](https://caniuse.com/)

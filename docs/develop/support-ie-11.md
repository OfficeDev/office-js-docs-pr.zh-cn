---
title: 支持 Internet Explorer 11
description: 了解如何在外接程序Internet Explorer 11 和 ES5 Javascript。
ms.date: 09/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3edab25361b8ababf8a004f25e8012ca23a085ab
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990521"
---
# <a name="support-internet-explorer-11"></a>支持 Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer外接程序Office中使用的内容**
>
> Microsoft 将终止对Internet Explorer的支持，但这不会显著影响Office外接程序。平台和 Office 版本（包括 Office 2019 的所有一次购买版本）的一些组合将继续使用 Internet Explorer 11 随附的 Webview 控件来托管外接程序，如[Office](../concepts/browsers-used-by-office-web-add-ins.md)外接程序使用的浏览器所说明。此外，提交到 AppSource 的加载项仍然需要支持这些组合Internet Explorer因此，这些组合对加载项[的支持也是必需的](/office/dev/store/submit-to-appsource-via-partner-center)。 有两 *个变化* ：
>
> - Office web 版中不再打开Internet Explorer。 因此，AppSource 不再使用作为浏览器Office web 版Internet Explorer加载项。 但 AppSource 仍测试使用 Office *版本的平台* 和桌面Internet Explorer。
> - Script Lab[工具](../overview/explore-with-script-lab.md)不再支持Internet Explorer。

Office外接程序是 Web 应用程序，当在 IFrame 上运行时，这些应用程序显示在 IFrame Office web 版。 Office加载项在 Mac 上的 Office 或 Windows Office浏览器控件中运行时显示。 嵌入式浏览器控件由操作系统或用户计算机上安装的浏览器提供。

如果计划通过 AppSource 销售加载项或计划支持较旧版本的 Windows 和 Office，加载项必须在基于 Internet Explorer 11 (IE11) 的可嵌入浏览器控件中运行。 有关使用基于 IE11 Windows和Office的浏览器控件的信息，请参阅 Office[外接程序使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!IMPORTANT]
> Internet Explorer 11 不支持某些 HTML5 功能，如媒体、录制和位置。 如果外接程序必须支持 Internet Explorer 11，则你无法使用这些功能。

Internet Explorer 11 不支持低于 ES5 的 JavaScript 版本。 如果要使用 ECMAScript 2015 或更高版本或 TypeScript 的语法和功能，有两个选项，如本文所述。 还可以结合这两种技术。

## <a name="use-a-transpiler"></a>使用转译器

可以使用 TypeScript 或新式 JavaScript 编写代码，然后在生成时将代码转换为 ES5 JavaScript。 生成的 ES5 文件是上传到加载项 Web 应用程序的文件。

有两种常用转译器。 两者都可以使用 TypeScript 或 ES5 后 JavaScript 的源文件。 它们还使用 React.jsx ( .tsx) 。

- [一些](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

有关在加载项项目中安装和配置转译器的信息，请参阅任一文档。 建议您使用任务运行程序（如 [Grunt](https://gruntjs.com/) 或 [WebPack）](https://webpack.js.org/) 来自动进行转换。 有关使用 tsc 的示例外接程序，请参阅 Office Microsoft 外接程序[Graph React。](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/auth/Office-Add-in-Microsoft-Graph-React) 有关使用分贝的示例，请参阅 Offline[存储 Add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/Excel.OfflineStorageAddin)。

> [!NOTE]
> 如果使用的不是Visual Studio (，Visual Studio Code) ，则 tsc 可能最易于使用。 可以使用 nuget 程序包安装对它的支持。 有关详细信息，请参阅[JavaScript and TypeScript in Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019)。 若要对任务Visual Studio，请创建生成脚本或使用 Visual Studio 中的任务运行程序资源管理器以及[WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner)任务运行程序或[NPM 任务运行程序等工具](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner)。

## <a name="use-a-polyfill"></a>使用填充

[填充是](https://en.wikipedia.org/wiki/Polyfill_(programming))早期版本的 JavaScript，与 JavaScript 的较新版本重复功能。 填充适用于不支持更高版本 JavaScript 的浏览器。 例如，字符串方法不是 JavaScript 的 ES5 版本的一部分，因此它不会在 `startsWith` 11 Internet Explorer中运行。 有一些用 ES5 编写的填充库定义了和实现 `startsWith` 一个方法。 我们建议使用 [core-js](https://github.com/zloirock/core-js) 填充库。

若要使用填充库，请像加载任何其他 JavaScript 文件或模块一样加载它。 例如，您可以使用外接程序主页 HTML 文件 (例如) 中的 标记，或者可以使用 JavaScript 文件 (例如) 中的 `<script>` `<script src="/js/core-js.js"></script>` `import` `import 'core-js';` 语句。 当 JavaScript 引擎看到类似 的方法时，它将首先查看语言中是否内置了该 `startsWith` 名称的方法。 如果存在，它将调用本机方法。 如果且仅在该方法未内置时，引擎将查找它的所有加载文件。 因此，填充版本不会在支持本机版本的浏览器中使用。

导入整个 core-js 库将导入所有 core-js 功能。 还可以仅导入加载项Office填充。 有关如何执行此操作的说明，请参阅[CommonJS API。](https://github.com/zloirock/core-js#commonjs-api) core-js 库具有所需的大部分填充。 core-js 文档的"缺少 [填充](https://github.com/zloirock/core-js#missing-polyfills) "部分详细介绍了一些例外情况。 例如，它不支持 `fetch` ，但可以使用 [fetch](https://github.com/github/fetch) 填充。

有关使用加载项示例core.js，请参阅 [Word 加载项 Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)。

## <a name="testing-an-add-in-on-internet-explorer"></a>在加载项上测试Internet Explorer

请参阅 [Internet Explorer 11 测试](../testing/ie-11-testing.md)。

## <a name="additional-resources"></a>其他资源

- [ECMAScript 6 兼容性表](https://kangax.github.io/compat-table/es6/)
- [我可以使用...HTML5、CSS3 等的支持表](https://caniuse.com/)

---
title: Internet Explorer 11 测试
description: 在 Office 11 上测试Internet Explorer加载项。
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8932545aa692073babeddb6ab22a213466a7c2ba
ms.sourcegitcommit: a3debae780126e03a1b566efdec4d8be83e405b8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/03/2021
ms.locfileid: "60809034"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>在 Office 11 上测试Internet Explorer加载项

> [!IMPORTANT]
> **Internet Explorer外接程序Office中使用的内容**
>
> Microsoft 将终止对Internet Explorer的支持，但这不会显著影响Office外接程序。平台和 Office 版本（包括 Office 2019 之间的一次购买版本）的一些组合将继续使用 Internet Explorer 11 随附的 Webview 控件来托管外接程序，如[Office](../concepts/browsers-used-by-office-web-add-ins.md)外接程序使用的浏览器所说明。此外，提交到[AppSource](/office/dev/store/submit-to-appsource-via-partner-center)的加载项仍然需要支持这些组合Internet Explorer，因此也支持这些组合。 有两 *个变化* ：
>
> - Office web 版中不再打开Internet Explorer。 因此，AppSource 不再使用 Office web 版 浏览器Internet Explorer测试加载项。 但是，AppSource 仍测试平台和 *Office的桌面* 版本的组合Internet Explorer。
> - Script Lab[工具](../overview/explore-with-script-lab.md)不再支持Internet Explorer。

如果计划通过 AppSource 销售加载项或计划支持较旧版本的 Windows 和 Office，加载项必须在基于 Internet Explorer 11 (IE11) 的可嵌入浏览器控件中运行。 可以使用命令行从外接程序使用的更现代运行时切换到 Internet Explorer 11 运行时进行此测试。 有关哪些版本的 Windows 和 Office使用 Internet Explorer 11 Web 视图控件，请参阅 Office[外接程序使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!IMPORTANT]
> Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。 如果要使用 ECMAScript 2015 或更高版本的语法和功能，有两个选项：
>
> - 在 ECMAScript 2015 (（也称为 ES6) 或更高版本 JavaScript）中编写代码，或在 TypeScript 中编写代码，然后使用编译器（如 [#A0](https://babeljs.io/) 或 [tsc）](https://www.typescriptlang.org/index.html)将代码编译为 ES5 JavaScript。
> - 在 ECMAScript 2015 或更高版本的 JavaScript[](https://en.wikipedia.org/wiki/Polyfill_(programming))中编写，但也加载填充库（如[core-js，](https://github.com/zloirock/core-js)它使 IE 能够运行代码）。
>
> 有关这些选项的详细信息，请参阅 Support [Internet Explorer 11](../develop/support-ie-11.md)。
>
> 此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。 若要了解更多信息，请参阅 [在运行时](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)确定加载项是否正在Internet Explorer。

> [!NOTE]
> Office web 版无法在 Internet Explorer 11 中打开，因此 (，也无需) 使用 Office web 版 测试Internet Explorer。

## <a name="switch-to-the-internet-explorer-11-webview"></a>切换到 Internet Explorer 11 Webview

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

有两种方法可以切换 web Internet Explorer视图。 可以在命令提示符中运行一个简单的命令，也可以安装默认Office使用Internet Explorer版本。 我们建议使用第一种方法。 但你应在以下方案中使用第二个。

- 您的项目是使用 Visual Studio IIS 开发的。 它不是基于node.js的。
- 你想要在测试中保持绝对可靠。
- 如果由于任何原因，命令行工具不起作用。

### <a name="switch-via-the-command-line"></a>通过命令行进行切换

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>安装使用Office版本的Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>另请参阅

* [测试和调试 Office 加载项](test-debug-office-add-ins.md)
* [旁加载 Office 外接程序进行测试](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [使用适用于 Internet Explorer 的开发人员工具调试加载项](debug-add-ins-using-f12-tools-ie.md)
* [从任务窗格附加调试器](attach-debugger-from-task-pane.md)

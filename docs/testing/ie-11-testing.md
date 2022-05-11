---
title: Internet Explorer 11 测试
description: 在 Internet Explorer 11 上测试Office加载项。
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: b8d027d4d583d42aa4efbe29e080afcd17297a74
ms.sourcegitcommit: fd04b41f513dbe9e623c212c1cbd877ae2285da0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/11/2022
ms.locfileid: "65313212"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>在 Internet Explorer 11 上测试Office加载项

> [!IMPORTANT]
> **Internet Explorer 仍在Office加载项中使用**
>
> 一些平台和Office版本的组合，包括到 2019 Office的一次性购买版本，仍然使用 Internet Explorer 11 附带的 Webview 控件来托管加载项，如[Office加载项使用的浏览器中](../concepts/browsers-used-by-office-web-add-ins.md)所述。建议 (但不需要) 继续支持这些组合（至少以最小方式）在 Internet Explorer Webview 中启动外接程序时为外接程序的用户提供正常故障消息。 请记住以下附加点：
>
> - Office web 版不再在 Internet Explorer 中打开。 因此，[AppSource](/office/dev/store/submit-to-appsource-via-partner-center) 不再使用 Internet Explorer 作为浏览器在Office web 版中测试加载项。
> - AppSource 仍在测试使用 Internet Explorer 的平台和Office *桌面* 版本的组合，但是仅当外接程序不支持 Internet Explorer 时才会发出警告;AppSource 不会拒绝该外接程序。
> - [Script Lab工具](../overview/explore-with-script-lab.md)不再支持 Internet Explorer。

如果计划支持较旧版本的Windows和Office，则加载项必须在基于 Internet Explorer 11 (IE11) 的可嵌入浏览器控件中工作。 可以使用命令行从外接程序使用的更新式运行时切换到 Internet Explorer 11 运行时进行此测试。 有关哪些版本的Windows和Office使用 Internet Explorer 11 Web 视图控件的信息，请参阅[Office加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!IMPORTANT]
> Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。 如果要使用 ECMAScript 2015 或更高版本的语法和功能，有两个选项：
>
> - 在 ECMAScript 2015 (也称为 ES6) 或更高版本 JavaScript 或 TypeScript 中编写代码，然后使用编译器（如 [babel](https://babeljs.io/) 或 [tsc](https://www.typescriptlang.org/index.html)）将代码编译到 ES5 JavaScript。
> - 在 ECMAScript 2015 或更高版本的 JavaScript 中编写，但还要加载 [一个多填充](https://en.wikipedia.org/wiki/Polyfill_(programming)) 库（如 [core-js](https://github.com/zloirock/core-js) ），使 IE 能够运行代码。
>
> 有关这些选项的详细信息，请参阅 [支持 Internet Explorer 11](../develop/support-ie-11.md)。
>
> 此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。 若要了解详细信息，请参阅 [运行时确定外接程序是否在 Internet Explorer 中运行](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)。

> [!NOTE]
> - Office web 版无法在 Internet Explorer 11 中打开，因此不能 (，也不需要) 使用 Internet Explorer 在Office web 版上测试加载项。
>
> - 必须关闭 Internet Explorer 的增强安全配置 (ESC) 才能使 Office Web 加载项正常工作。 如果在开发加载项时使用 Windows Server 计算机作为客户端，请注意 Windows Server 中会默认打开 ESC。

## <a name="switch-to-the-internet-explorer-11-webview"></a>切换到 Internet Explorer 11 Webview

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

有两种方法可以切换 Internet Explorer Webview。 可以在命令提示符下运行简单命令，也可以安装默认使用 Internet Explorer 的Office版本。 建议使用第一种方法。 但在以下方案中，应使用第二个。

- 你的项目是使用Visual Studio和 IIS 开发的。 它不是基于node.js的。
- 你希望在测试中绝对可靠。
- 如果出于任何原因，命令行工具不起作用。

### <a name="switch-via-the-command-line"></a>通过命令行切换

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>安装使用 Internet Explorer 的Office版本

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>另请参阅

* [测试和调试 Office 加载项](test-debug-office-add-ins.md)
* [旁加载 Office 外接程序进行测试](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [使用适用于 Internet Explorer 的开发人员工具调试加载项](debug-add-ins-using-f12-tools-ie.md)
* [从任务窗格附加调试器](attach-debugger-from-task-pane.md)

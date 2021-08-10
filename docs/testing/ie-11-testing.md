---
title: Internet Explorer 11 测试
description: 在 Office 11 上测试Internet Explorer加载项。
ms.date: 06/18/2021
localization_priority: Normal
ms.openlocfilehash: 8579a37f1ea48d511010b8c55cfe9fad5aa6b41acee85b1da426e25083287655
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57090124"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>在 Office 11 上测试Internet Explorer加载项

> [!IMPORTANT]
> **Internet Explorer外接程序Office中使用的内容**
>
> Microsoft 将终止对Internet Explorer的支持，但这不会显著Office外接程序。平台和 Office 版本（包括 Office 2019 的所有一次购买版本）的一些组合将继续使用 Internet Explorer 11 随附的 Webview 控件来托管外接程序，如[Office](../concepts/browsers-used-by-office-web-add-ins.md)外接程序使用的浏览器所说明。此外，提交到[AppSource](/office/dev/store/submit-to-appsource-via-partner-center)的加载项Internet Explorer支持这些组合，因此也支持这些组合。 有两 *个变化* ：
>
> - AppSource 不再使用作为浏览器Office web 版Internet Explorer加载项。 但 AppSource 仍测试使用 Office *版本的平台* 和桌面Internet Explorer。
> - 2021 Script Lab，Internet Explorer工具将停止工作。 [](../overview/explore-with-script-lab.md)

如果计划通过 AppSource 销售加载项或计划支持较旧版本的 Windows 和 Office，加载项必须在基于 Internet Explorer 11 (IE11) 的可嵌入浏览器控件中运行。 可以使用命令行从外接程序使用的更现代运行时切换到 Internet Explorer 11 运行时进行此测试。 有关哪些版本的 Windows 和 Office使用 Internet Explorer 11 Web 视图控件的信息，请参阅 Office [Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)使用的浏览器。

> [!IMPORTANT]
> Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。 如果要使用 ECMAScript 2015 或更高版本的语法和功能，有两个选项：
>
> - 在 ECMAScript 2015 (（也称为 ES6) 或更高版本 JavaScript）中编写代码，或在 TypeScript 中编写代码，然后使用编译器（如 [#A0](https://babeljs.io/) 或 [tsc）](https://www.typescriptlang.org/index.html)将代码编译为 ES5 JavaScript。
> - 在 ECMAScript 2015 或更高版本的 JavaScript[](https://en.wikipedia.org/wiki/Polyfill_(programming))中编写，但也加载填充库（如[core-js，](https://github.com/zloirock/core-js)它使 IE 能够运行代码）。
>
> 有关这些选项的详细信息，请参阅 Support [Internet Explorer 11](../develop/support-ie-11.md)。
>
> 此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。

> [!NOTE]
> 若要在 Internet Explorer 11 浏览器上测试外接程序，Office web 版中Internet Explorer并[旁加载外接程序](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)。

## <a name="prerequisites"></a>先决条件

- [Node.js](https://nodejs.org/)（最新的 [LTS](https://nodejs.org/about/releases) 版本）

这些说明假定你之前已经设置了 Yo Office生成器项目。 如果之前尚未这样做，请考虑阅读快速入门，例如适用于Excel[入门](../quickstarts/excel-quickstart-jquery.md)。

## <a name="switching-to-the-internet-explorer-11-webview"></a>切换到 Internet Explorer 11 Webview

1. 创建 Yo Office生成器项目。 选择哪种项目并不重要，此工具将用于所有项目类型。

    > [!NOTE]
    > 如果您有一个现有项目，并且想要在不创建新项目的情况下添加此工具，请跳过此步骤并移至下一步。 

1. 在项目的根文件夹中，在命令行中运行以下命令。 此示例假定项目的清单文件位于根中。 如果不是，请指定清单文件的相对路径。 您应该在命令行中看到一条消息，指出 Web 视图类型现在设置为 IE。

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> 虽然不需要使用此命令，但它应有助于调试与 11 运行时Internet Explorer大多数问题。 为提供完整的稳定性，应测试使用具有 Windows 7、8.1 和 10 的各种版本以及不同版本的 Office 的计算机。 有关详细信息，请参阅Office[外接程序](../concepts/browsers-used-by-office-web-add-ins.md)使用的浏览器和如何还原到早期版本[Office。](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841)

### <a name="command-options"></a>命令选项

该命令 `office-addin-dev-settings webview` 还可以将多个运行时用作参数：

- ie
- edge
- default

## <a name="see-also"></a>另请参阅

* [测试和调试 Office 加载项](test-debug-office-add-ins.md)
* [旁加载 Office 外接程序进行测试](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [使用 Windows 10 上的开发人员工具调试加载项](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [从任务窗格附加调试器](attach-debugger-from-task-pane.md)

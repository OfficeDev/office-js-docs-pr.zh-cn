---
title: 调试 Office 加载项
description: 查找开发环境的 Office 加载项调试指南。
ms.date: 07/11/2022
ms.localizationpriority: high
ms.openlocfilehash: e8f4270a133e068333703796c10b091bae31ed0a
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797482"
---
# <a name="overview-of-debugging-office-add-ins"></a>调试 Office 加载项概述

调试 Office 加载项实质上与调试任何 Web 应用程序相同。 但是，一组工具不适用于所有加载项开发人员。 这是因为加载项可以在不同的操作系统上开发并跨平台运行。 本文可帮助你找到开发环境的详细调试指南。

> [!TIP]
> 本文关注的是狭义上的调试，即设置断点和单步执行代码。 有关测试和故障排除的指南，请从 [测试 Office 加载项](test-debug-office-add-ins.md) 和 [使用 Office 加载项排查开发错误](troubleshoot-development-errors.md) 开始。

> [!NOTE]
> 尽管应在要支持的所有平台上 *测试* 加载项，但在不同于开发计算机的环境中，你只需要进行 *调试*。 因此，本文使用“开发计算机”和“你的开发环境”来表示要进行调试的环境。 如果代码中的问题仅发生在开发计算机以外的平台上，并且需要设置断点或单步执行代码来解决该问题，则进行调试的环境并不是你的开发环境。

## <a name="server-side-or-client-side"></a>服务器端还是客户端？

调试 Office 加载项的服务器端代码与调试任何 Web 应用程序的服务器端相同。 请参阅 IDE 或其他工具的调试说明。 下面是一些最常用工具的示例。

- [在 Visual Studio 中调试 ASP.NET 或 ASP.NET Core 应用](/visualstudio/debugger/how-to-enable-debugging-for-aspnet-applications)
- [调试 Express](https://expressjs.com/en/guide/debugging.html)
- [Node.js 调试指南](https://nodejs.org/en/docs/guides/debugging-getting-started/)
- [VS Code 中的 Node.js 调试](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Webpack 调试](https://webpack.js.org/contribute/debugging/)

本文的其余部分仅涉及调试客户端 JavaScript（可从 TypeScript 转译）。

## <a name="special-cases"></a>特殊情况

在某些特殊情况下，对于给定的平台、Office 应用程序和开发环境组合，调试过程不同于正常情况。 如果要调试这些特殊情况中的任何一种，请使用本部分中的链接来查找适当的指南。 否则，请继续使用 [常规指南](#general-guidance)。

- **调试 `Office.initialize` 或 `Office.onReady` 方法**：[调试 initialize 和 onReady 方法](debug-initialize-onready.md)。
- **在 _非共享_ 运行时中调试 Excel 自定义函数**：[非共享运行时中的自定义函数调试](../excel/custom-functions-debugging.md)。
- **在 _非共享_ 运行时中调试 [函数命令](../design/add-in-commands.md#types-of-add-in-commands)**： 
    - Windows 开发计算机上的 Outlook 加载项：[在 Outlook 加载项中调试函数命令](../outlook/debug-ui-less.md) 
    - Mac 开发计算机上的其他 Office 应用程序加载项或 Outlook：[使用非共享运行时调试函数命令](debug-function-command.md)。
- **调试基于事件的 Outlook 加载项**：[调试基于事件的 Outlook 加载项](../outlook/debug-autolaunch.md)。 
 
## <a name="general-guidance"></a>一般指导

如果要查找有关调试客户端代码的指南，则第一个变量是开发计算机的操作系统。

- [Windows](#debug-on-windows)
- [Mac](#debug-on-mac)
- [Linux 或其他 Unix 变体](#debug-on-linux)

### <a name="debug-on-windows"></a>在 Windows 上调试

下面提供了有关在 Windows 上进行调试的常规指南。 在 Windows 上调试取决于 IDE。

- **Visual Studio**：使用浏览器的 F12 工具进行调试。 请参阅 [在 Visual Studio 中调试 Office 加载项](../develop/debug-office-add-ins-in-visual-studio.md)。
- **Visual Studio Code**：使用 [适用于 Visual Studio Code 的加载项调试器扩展](debug-with-vs-extension.md) 进行调试。
- **任何其他 IDE**（或者你不想在 IDE 内部进行调试）：使用与加载项在开发计算机上使用的浏览器运行时关联的开发人员工具。请参阅下列文档之一：

    - [使用适用于 Internet Explorer 的开发人员工具调试加载项](debug-add-ins-using-f12-tools-ie.md)
    - [使用旧版 Edge 开发人员工具调试加载项](debug-add-ins-using-devtools-edge-legacy.md)
    - [使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项](debug-add-ins-using-devtools-edge-chromium.md)

有关正在使用哪个浏览器运行时的信息，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

### <a name="debug-on-mac"></a>在 Mac 上调试

下面提供了有关在 Mac 上进行调试的常规指南。

- 如果使用 Visual Studio Code，请使用 [适用于 Visual Studio Code 的加载项调试器扩展](debug-with-vs-extension.md) 进行调试。
- 对于任何其他 IDE，请使用 Safari Web 检查器。 说明位于 [在 Mac 上调试 Office 加载项](debug-office-add-ins-on-ipad-and-mac.md) 中。


### <a name="debug-on-linux"></a>在 Linux 上调试

没有适用于 Linux 的 Office 桌面版本，因此需要 [将加载项旁加载到 Office 网页版](sideload-office-add-ins-for-testing.md)才能对其进行测试和调试。调试指南位于[在 Office 网页版中调试加载项](debug-add-ins-in-office-online.md)中。

> [!NOTE]
> 除可以确保所有加载项用户都将从 Linux 计算机通过 Office 网页版访问加载项的少数情况以外，我们不建议在 Linux 计算机上开发 Office 加载项。

## <a name="debug-add-ins-in-staging-or-production"></a>在暂存或生产中调试加载项

要调试已在暂存或生产中的加载项，请从加载项的 UI 附加调试程序。 要了解说明，请参阅 [从任务窗格中附加调试程序](attach-debugger-from-task-pane.md)。
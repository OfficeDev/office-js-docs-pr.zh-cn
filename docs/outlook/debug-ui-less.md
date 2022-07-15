---
title: Outlook 外接程序中的调试函数命令
description: 了解如何在 Outlook 加载项中调试函数命令。
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6189824fd526d48321b355c9b306fa5ef732f411
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797587"
---
# <a name="debug-function-commands-in-outlook-add-ins"></a>Outlook 外接程序中的调试函数命令

> [!NOTE]
> 本文中的技术只能在 Windows 开发计算机上使用。 如果要在 Mac 上进行开发，请参阅 [“调试函数”命令](../testing/debug-function-command.md)。

本文介绍如何在 Visual Studio Code 中使用 Office 加载项调试器扩展来调试[函数命令](add-in-commands-for-outlook.md#run-a-function-command)。 函数命令是通过功能区中的加载项命令按钮启动的。 有关外接程序命令的详细信息，请参阅 [Outlook 的外接程序命令](add-in-commands-for-outlook.md)。

本文假定你已有要调试的加载项项目。 若要使用函数命令创建加载项以练习调试，请按照教程中的步骤操作 [：生成消息撰写 Outlook 加载项](../tutorials/outlook-tutorial.md)。

## <a name="mark-your-add-in-for-debugging"></a>标记加载项以进行调试

如果使用 [Office 外接程序的 Yeoman 生成器](../develop/yeoman-generator-overview.md) 创建外接程序项目，请跳到“配置”，并在本文后面 [运行调试器](#configure-and-run-the-debugger) 部分。 运行 `npm start` 以生成外接程序并启动本地服务器时，该命令还会设置 `UseDirectDebugger` 注册表项的 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` 值，以标记用于调试的外接程序。

否则，如果使用其他工具创建外接程序，请执行以下步骤。

1. 导航到 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` 注册表项。 替换 `[Add-in ID]` 为 **\<Id\>** 加载项清单中的清单。

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. 将键的 `UseDirectDebugger` 值设置为 `1`.

## <a name="configure-and-run-the-debugger"></a>配置并运行调试器

在加载项上启用调试后，便可以配置并运行调试器了。 有关如何执行此操作的说明，请选择下列适用于 Webview 控件的选项之一。 有关如何确定开发计算机上使用的 Webview 控件的信息，请参阅 [Office 外接程序使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

- 如果外接程序在 Edge 旧版 (EdgeHTML) 的嵌入式 Web 视图控件中运行，请参阅[适用于 Visual Studio Code 的 Microsoft Office 外接程序调试器扩展](../testing/debug-with-vs-extension.md)。

- 如果外接程序在 Microsoft Edge Chromium (WebView2) 的嵌入式 Web 视图控件中运行，请参阅[使用基于 Visual Studio Code 和 Microsoft Edge WebView2 的 Windows 上的调试加载项 (Chromium) ](../testing/debug-desktop-using-edge-chromium.md)。

## <a name="see-also"></a>另请参阅

- [适用于 Outlook 的外接程序命令](add-in-commands-for-outlook.md)
- [调试 Office 加载项概述](../testing/debug-add-ins-overview.md)
- [调试基于事件的 Outlook 加载项](debug-autolaunch.md)

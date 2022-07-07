---
title: 调试无 UI Outlook 加载项
description: 了解如何调试无 UI 的 Outlook 外接程序。
ms.topic: article
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: e46bdf15172f5224995b17c39df4ba60ca6380ad
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660205"
---
# <a name="debug-your-ui-less-outlook-add-in"></a>调试无 UI Outlook 加载项

本文介绍如何在 Visual Studio Code 中使用 Office 加载项调试器扩展来调试[无 UI 的 Outlook 外接程序](add-in-commands-for-outlook.md#executing-a-javascript-function)。无 UI 加载项操作是通过功能区中的加载项命令按钮启动的。 有关外接程序命令的详细信息，请参阅 [Outlook 的外接程序命令](add-in-commands-for-outlook.md)。

本文假定你已有要调试的加载项项目。 若要创建无 UI 加载项来练习调试，请按照 [教程中的步骤操作：生成消息撰写 Outlook 加载项](../tutorials/outlook-tutorial.md)。

## <a name="mark-your-add-in-for-debugging"></a>标记加载项以进行调试

如果使用 [Office 外接程序的 Yeoman 生成器](../develop/yeoman-generator-overview.md) 创建外接程序项目，请跳到“配置”，并在本文后面 [运行调试器](#configure-and-run-the-debugger) 部分。 运行 `npm start` 以生成外接程序并启动本地服务器时，该命令还会设置 `UseDirectDebugger` 注册表项的 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` 值，以标记用于调试的外接程序。

否则，如果使用其他工具创建外接程序，请执行以下步骤。

1. 导航到 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` 注册表项。 替换 `[Add-in ID]` 为 **\<Id\>** 加载项清单中的清单。

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. 将键的 `UseDirectDebugger` 值设置为 `1`.

## <a name="configure-and-run-the-debugger"></a>配置并运行调试器

在加载项上启用调试后，便可以配置并运行调试器了。 有关如何执行此操作的说明，请选择适用于运行时的以下选项之一。

- 如果加载项在 WebView 运行时中运行，请参阅 [Microsoft Office 加载项调试器扩展以Visual Studio Code](../testing/debug-with-vs-extension.md)。

- 如果外接程序在 Microsoft Edge Chromium WebView2 运行时中运行，请参阅[使用基于 Visual Studio Code 和 Microsoft Edge WebView2 的 Windows 上的调试加载项 (Chromium) ](../testing/debug-desktop-using-edge-chromium.md)。

## <a name="see-also"></a>另请参阅

- [适用于 Outlook 的外接程序命令](add-in-commands-for-outlook.md)
- [调试 Office 加载项概述](../testing/debug-add-ins-overview.md)
- [调试基于事件的 Outlook 加载项](debug-autolaunch.md)

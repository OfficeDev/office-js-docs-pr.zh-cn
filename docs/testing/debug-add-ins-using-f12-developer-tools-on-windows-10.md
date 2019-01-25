---
title: 在 Windows 10 上使用 F12 开发人员工具调试外接程序
description: ''
ms.date: 10/16/2018
localization_priority: Priority
ms.openlocfilehash: e2378a0449ea33551051b9c3788b84b23a51feb8
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386902"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>在 Windows 10 上使用 F12 开发人员工具调试外接程序

Windows 10 中随附的 F12 开发人员工具可帮助您调试、测试和加速您的网页。 如果您未使用 IDE（如 Visual Studio），或者如果您需要调查在 IDE 外部运行外接程序时出现的问题，您还可以使用该工具开发和调试您的 Office 外接程序。 本文介绍如何在 Windows 10 中使用 F12 开发人员工具中的调试器工具来测试你的 Office 加载项。

> [!NOTE]
> 本文中的说明不能用于调试使用 Execute 函数的 Outlook 加载项。 若要调试使用 Execute 函数的 Outlook 加载项，我们建议你在脚本模式下附加到 Visual Studio 或附加到某些其他脚本调试器。

## <a name="prerequisites"></a>先决条件

您需要安装以下软件：

- Windows 10 随附的 F12 开发人员工具 
    
- 托管您的外接程序的 Office 客户端应用程序。  
    
- 您的外接程序。  

## <a name="using-the-debugger"></a>使用调试器

本文介绍了如何在 Windows 10 上使用 F12 开发人员工具中的调试器工具，测试 Office 加载项。可以测试从 AppSource 获取的加载项，也可以测试从其他位置添加的加载项。F12 工具在单独的窗口中显示，并不使用 Visual Studio。 运行加载项后，可以启动 F12 开发人员工具。 F12 工具显示在单独的窗口中，并不使用 Visual Studio。

> [!NOTE]
> 调试器属于 Windows 10 和 Internet Explorer 上的 F12 开发人员工具。旧版 Windows 不包含调试器。 

此示例使用 Word 和从 AppSource 获取的免费加载项。

1. 打开 Word 并选择空白文档。 
    
2. 在“**插入**”选项卡上的“加载项”组中，依次选择“**存储**”和 **QR4Office** 加载项。 （你可以从应用商店或加载项目录中加载任何加载项。）
    
3. 启动与 Office 版本相对应的 F12 开发工具：
    
   - 对于 32 位版 Office，请使用 C:\Windows\System32\F12\IEChooser.exe
    
   - 对于 64 位版 Office，请使用 C:\Windows\SysWOW64\F12\IEChooser.exe
    
   当你启动 IEChooser 时，一个单独的窗口（名为“选择要调试的目标”）显示要调试的可能的应用程序。 选择你感兴趣的应用程序。 如果你正在编写自己的加载项，请选择你已在其中部署加载项的网站，这可能是本地主机 URL。 
    
   例如，选择 **home.html**。 
    
   ![IEChooser 屏幕，指向圈出的加载项](../images/choose-target-to-debug.png)

4. 在 F12 窗口中，选择你想要调试的文件。
    
   若要在 F12 窗口中选择文件，请选择“**脚本**”（左）窗格上方的文件夹图标。 从下拉列表中显示的可用文件列表中，选择 **Home.js**。
    
5. 设置断点。
    
   若要在 **Home.js** 中设置断点，请选择第 144 行，它位于 `textChanged` 函数中。 你将在该行左侧和“调用堆栈和断点”（右下角）窗格中的对应行左侧看到一个红点。 有关设置断点的其他方法，请参阅[使用调试器检查正在运行的 JavaScript](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))。 
    
   ![断点位于 home.js 文件中的调试程序](../images/debugger-home-js-02.png)

6. 运行加载项，以触发断点。
    
   在 Word 中，选择 **QR4Office** 窗格上部的 URL 文本框，然后尝试输入一些文本。 在调试器的“**调用堆栈和断点**”窗格中，你将看到该断点已触发，并显示了各种信息。 你可能需要刷新调试器以查看结果。
    
   ![调试器，包含已触发的断点生成的结果](../images/debugger-home-js-01.png)


## <a name="see-also"></a>另请参阅

- [使用调试器检查正在运行的 JavaScript](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [使用 F12 开发人员工具](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))

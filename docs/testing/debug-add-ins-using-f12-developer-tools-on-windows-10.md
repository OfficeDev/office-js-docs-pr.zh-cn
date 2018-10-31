---
title: 在 Windows 10 上使用 F12 开发人员工具调试外接程序
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 3df245fcd651ec227e0a32d53da186ee332beb8f
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579840"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>在 Windows 10 上使用 F12 开发人员工具调试外接程序

Windows 10 中随附的 F12 开发人员工具可帮助您调试、测试和加速您的网页。如果您未使用 IDE（如 Visual Studio），或者如果您需要调查在 IDE 外部运行外接程序时出现的问题，您还可以使用它们来开发和调试 Office 外接程序。本文说明了如何在 Windows 10 中使用来自 F12 开发人员工具的调试工具来测试您的 Office 外接程序。

> [!NOTE]
> 本文中的说明不适用于调试使用 Execute 函数的 Outlook 加载项。 若要调试使用 Execute 函数的 Outlook 加载项，建议附加到脚本模式中的 Visual Studio 或某些其他脚本调试程序。

## <a name="prerequisites"></a>先决条件

您需要安装以下软件：

- Windows 10 随附的 F12 开发人员工具。 
    
- 您的外接程序的宿主 Office 客户端应用程序。 
    
- 您的外接程序。 

## <a name="using-the-debugger"></a>使用调试器

可以使用在 Windows 10 从 F12 开发人员工具调试工具测试来自 AppSource 的外接程序或从其他位置添加的外接程序。 运行外接程序后，可以启动 F12 开发人员工具。 F12 工具在一个单独的窗口中显示，并且不使用 Visual Studio。

> [!NOTE]
> 调试器属于 Windows 10 和 Internet Explorer 上的 F12 开发人员工具。旧版 Windows 不包含调试器。 

此示例使用 Word 和从 AppSource 获取的免费外接程序。

1. 打开 Word 并选择空白文档。 
    
2. 在**插入**选项卡的”外接程序“组中，选择**存储** ，然后选择 **QR4Office** 外接程序。 （您可以从应用商店或外接程序目录中加载任何外接程序。）
    
3. 启动与 Office 版本相对应的 F12 开发工具：
    
   - 对于 32 位版 Office，请使用 C:\Windows\System32\F12\IEChooser.exe
    
   - 对于 64 位版 Office，请使用 C:\Windows\SysWOW64\F12\IEChooser.exe
    
   当你启动 IEChooser 时，一个单独的窗口（名为“选择要调试的目标”）显示可能要调试的应用程序。 选择你感兴趣的应用程序。 如果你正在编写自己的外接程序，请选择你已在其中部署外接程序的网站，这可能是 localhost URL。 
    
   例如，选择 ** home.html**。 
    
   ![IEChooser 界面，指向气泡加载项](../images/choose-target-to-debug.png)

4. 在 F12 窗口中，选择您想要调试的文件。
    
   若要选择 F12 窗口中的文件，请选择**脚本**（左）窗格上方的文件夹图标。 从下拉列表中显示的可用文件列表中，选择 **Home.js**。
    
5. 设置断点。
    
   若要在 **home.js** 中设置断点，请选择第 144 行，它位于 `textChanged` 函数中。 你将在该行左侧和**调用堆栈和断点**（右下角）窗格中的对应行左侧看到一个红点。 有关设置断点的其他方法，请参阅[使用调试器检查正在运行的 JavaScript](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))。 
    
   ![在 home.js 文件中包含断点的调试程序](../images/debugger-home-js-02.png)

6. 运行外接程序以触发断点。
    
   在 Word 中，在 **QR4Office** 窗格的上部选择 URL 文本框，并尝试输入一些文本。 在调试器的**调用堆栈和断点**窗格中，你将看到该断点已触发，并显示了各种信息。 你可能需要刷新调试器工具以查看结果。
    
   ![包含已触发断点生成结果的调试器](../images/debugger-home-js-01.png)


## <a name="see-also"></a>另请参阅

- [使用调试器检查正在运行的 JavaScript](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [使用 F12 开发人员工具](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))

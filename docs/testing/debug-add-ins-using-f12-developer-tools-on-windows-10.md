---
title: 在 Windows 10 上使用 F12 开发人员工具调试外接程序
description: ''
ms.date: 01/23/2018
---

# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>在 Windows 10 上使用 F12 开发人员工具调试外接程序

Windows 10 中随附的 F12 开发人员工具可帮助您调试、测试和加速您的网页。如果您未使用 IDE（如 Visual Studio），或者如果您需要调查在 IDE 外部运行外接程序时出现的问题，您还可以使用该工具开发和调试您的 Office 外接程序。运行外接程序后，可以启动 F12 开发人员工具。

本文介绍了如何在 Windows 10 上使用 F12 开发人员工具中的调试器工具，测试 Office 加载项。可以测试从 AppSource 获取的加载项，也可以测试从其他位置添加的加载项。F12 工具在单独的窗口中显示，并不使用 Visual Studio。

> [!NOTE]
> 调试器属于 Windows 10 和 Internet Explorer 上的 F12 开发人员工具。旧版 Windows 不包含调试器。 

## <a name="prerequisites"></a>先决条件

您需要安装以下软件：

- Windows 10 随附的 F12 开发人员工具 
    
- 托管您的外接程序的 Office 客户端应用程序。  
    
- 您的外接程序。  

## <a name="using-the-debugger"></a>使用调试器

此示例使用 Word 和从 AppSource 获取的免费加载项。

1. 打开 Word 并选择空白文档。 
    
2. 在“插入”****选项卡上的“加载项”组中，依次选择“Microsoft Store”****和 QR4Office 加载项。（可以从 Microsoft Store 或加载项目录中加载任何加载项。）
    
3. 启动与 Office 版本相对应的 F12 开发工具：
    
   - 对于 32 位版 Office，请使用 C:\Windows\System32\F12\F12Chooser.exe
    
   - 对于 64 位版 Office，请使用 C:\Windows\SysWOW64\F12\F12Chooser.exe
    
   当你启动 F12Chooser 时，一个单独的窗口（名为“选择要调试的目标”）显示要调试的可能的应用程序。选择你感兴趣的应用程序。如果你正在编写自己的外接程序，请选择你已在其中部署外接程序的网站，这可能是本地主机 URL。 
    
   例如，选择“home.html”****。 
    
   ![F12Chooser 屏幕，指向圈出的加载项](../images/choose-target-to-debug.png)

4. 在 F12 窗口中，选择您想要调试的文件。
    
   若要选择文件，请选择“**脚本**”（左）窗格上方的文件夹图标。下拉列表显示了可用文件。选择 home.js。
    
5. 设置断点。
    
   To set the breakpoint in home.js, choose line 144, which is in the  _textChanged_ function. You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane. For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx). 
    
   ![断点位于 home.js 文件中的调试程序](../images/debugger-home-js-02.png)

6. 运行加载项，以触发断点。
    
   选择 QR4Office 窗格上半部分中的 URL 文本框，以更改文本。在“调试器”的“调用堆栈和断点”****窗格中，将看到断点已触发，以及显示的各种信息。建议刷新 F12 工具来查看结果。
    
   ![调试器，包含已触发的断点生成的结果](../images/debugger-home-js-01.png)


## <a name="see-also"></a>另请参阅

- [使用调试器检查正在运行的 JavaScript](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx)
- 
  [使用 F12 开发人员工具](https://msdn.microsoft.com/zh-cn/library/bg182326%28v=vs.85%29.aspx)
    

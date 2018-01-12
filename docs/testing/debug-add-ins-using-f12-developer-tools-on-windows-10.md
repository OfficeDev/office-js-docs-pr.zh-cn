
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>在 Windows 10 上使用 F12 开发人员工具调试外接程序

Windows 10 中随附的 F12 开发人员工具可帮助您调试、测试和加速您的网页。如果您未使用 IDE（如 Visual Studio），或者如果您需要调查在 IDE 外部运行外接程序时出现的问题，您还可以使用该工具开发和调试您的 Office 外接程序。运行外接程序后，可以启动 F12 开发人员工具。

本文介绍如何在 Windows 10 中使用 F12 开发人员工具中的调试器工具来测试您的 Office 外接程序。您可以从 Office 应用商店中或从其他位置添加的外接程序测试外接程序。F12 工具显示在其各自的窗口中，并不使用 Visual Studio。

 >**注意** 调试器是 Windows 10 上 F12 开发人员工具和 Internet Explorer 的一部分。较低版本的 Windows 不包含调试器。 


### <a name="prerequisites"></a>先决条件

您需要安装以下软件：


- Windows 10 随附的 F12 开发人员工具 
    
- 托管您的外接程序的 Office 客户端应用程序。  
    
- 您的外接程序。  
    
### <a name="using-the-debugger"></a>使用调试器

此示例使用 Word 和 Office 应用商店的免费外接程序。

1. 打开 Word 并选择一个空白文档。  
    
2. 在“**插入**”选项卡上的“外接程序”组中，存储并选择 QR4Office 外接程序。（您可以从应用商店或外接程序目录中加载任何外接程序。）
    
3. 启动与您的 Office 版本相对应的 F12 开发工具：
    
      - 对于 32 位版本的 Office，使用 C:\Windows\System32\F12\F12Chooser.exe
    
  - 对于 64 位版本的 Office，使用 C:\Windows\SysWOW64\F12\F12Chooser.exe
    

    当你启动 F12Chooser 时，一个单独的窗口（名为“选择要调试的目标”）显示要调试的可能的应用程序。选择你感兴趣的应用程序。如果你正在编写自己的外接程序，请选择你已在其中部署外接程序的网站，这可能是本地主机 URL。 
    
    例如，选择 **home.html**。 
    
    ![F12Chooser 屏幕，指向气泡外接程序](../../images/4f8823a3-595a-4657-83ac-8b235a7ba087.png)

4. 在 F12 窗口中，选择您想要调试的文件。
    
    若要选择文件，请选择“**脚本**”（左）窗格上方的文件夹图标。下拉列表显示了可用文件。选择 home.js。
    
5. 设置断点。
    
    To set the breakpoint in home.js, choose line 144, which is in the  _textChanged_ function. You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane. For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx). 
    
    ![断点位于 home.js 文件中的调试程序](../../images/e3cbc7ca-8b21-4ebb-b7a1-93e2364f1d16.png)

6. 运行外接程序以触发断点。
    
    选择 QR4Office 窗格上半部分中的 URL 文本框更改文本。在调试器的“**调用堆栈和断点**”窗格中，你将看到该断点已触发，并显示了各种信息。你可能需要刷新 F12 工具以查看结果。
    
    ![结果来自触发断点的调试程序](../../images/e0bcd036-91ce-4509-ae98-6c10b593d61b.png)


## <a name="additional-resources"></a>其他资源



- [使用调试器检查正在运行的 JavaScript](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx)
    
- [使用 F12 开发人员工具](https://msdn.microsoft.com/en-us/library/bg182326%28v=vs.85%29.aspx)
    

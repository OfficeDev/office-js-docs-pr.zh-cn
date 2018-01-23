
# <a name="debug-add-ins-in-office-online"></a>调试 Office Online 中的外接程序


您可以在并非运行 Windows 或 Office 2013 或 Office 2016 桌面客户端的计算机上构建和调试外接程序，例如，如果您正在使用 Mac 进行开发。本文介绍如何使用 Office Online 测试和调试您的外接程序。 

若要开始，请执行以下操作：


- 获取 Office 365 开发人员帐户（如果还没有）或者获取 SharePoint 站点的访问权限。
    
     >**注意**  若要注册免费的 Office 365 开发人员帐户，请加入我们的 [Office 365 开发人员计划](https://dev.office.com/devprogram)。
     
- 在 Office 365 (SharePoint Online) 上设置外接程序目录。 外接程序目录是 SharePoint Online 中的专用网站集，它托管 Office 外接程序的文档库。 如果你有自己的 SharePoint 站点，则可以设置外接程序目录文档库。 有关详细信息，请参阅 [向 SharePoint 上的外接程序目录发布任务窗格和内容外接程序](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a>从 Excel Online 或 Word Online 调试外接程序

要使用 Office Online 调试您的外接程序，请执行以下操作：


1. 将外接程序部署到支持 SSL 的服务器上。
    
     >**注意：**我们建议使用 [Yeoman 生成器](https://github.com/OfficeDev/generator-office)以创建和托管外接程序。
     
2. 在你的 [外接程序清单文件](../overview/add-in-manifests.md)中，将 **SourceLocation** 元素值更新为包括绝对 URI，而非相对 URI。 例如：
    
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. 将清单上传到 SharePoint 上的外接程序目录中的 Office 外接程序库。
    
4. 从 Office 365 中的应用程序启动程序启动 Excel Online 或 Word Online，并打开一个新文档。
    
5. 在“插入”选项卡上，选择“**我的外接程序**”或“**Office 外接程序**”以插入你的外接程序并在应用中对其测试。
    
6. 使用您最喜欢的浏览器工具调试程序调试您的外接程序。
    
    下面介绍了一些你可能会在调试过程中遇到的问题：
    
  - 您看到的一些 JavaScript 错误可能源自 Office Online。
    
  - 浏览器可能会显示无效证书错误，您需绕过此错误。
    
  - 如果您在代码中设置了断点，Office Online 可能会抛出一个错误，指示无法保存断点。
    

## <a name="additional-resources"></a>其他资源


- [开发 Office 外接程序的最佳做法](../overview/add-in-development-best-practices.md)
    
- [提交到 Office 应用商店的应用和外接程序的验证策略（版本 1.9）](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)
    
- [创建有效的 Office 应用商店应用和外接程序](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)
    
- [解决 Office 外接程序中的用户错误](../testing/testing-and-troubleshooting.md)
    

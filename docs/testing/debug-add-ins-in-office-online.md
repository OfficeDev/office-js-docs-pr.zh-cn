---
title: 在 Office Online 中调试加载项
description: ''
ms.date: 01/23/2018
---

# <a name="debug-add-ins-in-office-online"></a>在 Office Online 中调试加载项


您可以在并非运行 Windows 或 Office 2013 或 Office 2016 桌面客户端的计算机上构建和调试外接程序，例如，如果您正在使用 Mac 进行开发。本文介绍如何使用 Office Online 测试和调试您的外接程序。 

## <a name="prerequisites"></a>先决条件

首先，请执行以下操作：

- 获取 Office 365 开发人员帐户（如果还没有的话），或获取对 SharePoint 网站的访问权限。
    
  > [!NOTE]
  > 若要注册免费的 Office 365 开发人员帐户，请加入 [Office 365 开发人员计划](https://dev.office.com/devprogram)。
     
- 在 Office 365 (SharePoint Online) 上设置外接程序目录。外接程序目录是 SharePoint Online 中的专用网站集，它托管 Office 外接程序的文档库。如果你有自己的 SharePoint 站点，则可以设置外接程序目录文档库。有关详细信息，请参阅 [向 SharePoint 上的外接程序目录发布任务窗格和内容外接程序](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a>从 Excel Online 或 Word Online 调试外接程序

要使用 Office Online 调试您的外接程序，请执行以下操作：

1. 将加载项部署到支持 SSL 的服务器上。
    
    > [!NOTE]
    > 建议使用 [Yeoman 生成器](https://github.com/OfficeDev/generator-office)创建和托管加载项。
     
2. 在[加载项清单文件](../develop/add-in-manifests.md)中，将 **SourceLocation** 元素值更新为包括绝对 URI，而不是相对 URI。例如：
      
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. 将清单上传到 SharePoint 上加载项目录中的“Office 加载项”库。
    
4. 从 Office 365 中的应用程序启动程序启动 Excel Online 或 Word Online，并打开一个新文档。
    
5. 在“插入”选项卡上，选择“**我的外接程序**”或“**Office 外接程序**”以插入你的外接程序并在应用中对其测试。
    
6. 使用常用浏览器工具调试器调试加载项。

## <a name="potential-issues"></a>潜在问题    

下面介绍了一些在调试过程中可能会遇到的问题：
    
- 您看到的一些 JavaScript 错误可能源自 Office Online。
      
- 浏览器可能会显示无效证书错误，您需绕过此错误。
      
- 如果在代码中设置了断点，Office Online 可能会抛出错误，指示无法保存。

## <a name="see-also"></a>另请参阅

- [Office 加载项开发最佳做法](../concepts/add-in-development-best-practices.md)
- 
  [AppSource 验证策略](https://docs.microsoft.com/zh-cn/office/dev/store/validation-policies)  
- 
  [创建有效的 AppSource 应用和加载项](https://docs.microsoft.com/zh-cn/office/dev/store/create-effective-office-store-listings)  
- [排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)
    

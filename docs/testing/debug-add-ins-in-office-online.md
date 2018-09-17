---
title: 在 Office Online 中调试加载项
description: 如何使用 Office Online 测试和调试加载项。
ms.date: 03/14/2018
ms.openlocfilehash: ee458352c78a3bb7828e66df9fcde12958f3df93
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945762"
---
# <a name="debug-add-ins-in-office-online"></a>在 Office Online 中调试加载项


您可以生成和调试未运行 Windows 或 Office 桌面客户端计算机上加载&mdash;例如，如果您正在开发 mac。 如何使用 Office Online 测试和调试加载项。 

## <a name="prerequisites"></a>先决条件

首先，请执行以下操作：

- 获取 Office 365 开发人员帐户（如果还没有的话），或获取对 SharePoint 网站的访问权限。
    
  > [!NOTE]
  > 若要注册免费 Office 365 开发人员订阅，请加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)。 请参阅 [Office 365 开发人员计划文档](https://docs.microsoft.com/office/developer-program/office-365-developer-program)，逐步了解如何加入 Office 365 开发人员计划并注册和配置订阅。
     
- 对 Office 365 (SharePoint Online) 设置加载项目录。加载项目录是 SharePoint Online 中的专用网站集，用于托管 Office 加载项的文档库。如果有自己的 SharePoint 网站，可以设置加载项目录文档库。有关详细信息，请参阅[向 SharePoint 上的加载项目录发布任务窗格和内容加载项](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a>通过 Excel Online 或 Word Online 调试加载项

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
- [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)  
- [创建有效的 AppSource 应用和加载项](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)  
- [排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)
    

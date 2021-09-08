---
title: 在 Office 网页版中调试加载项
description: 如何使用 Office 网页版来测试和调试加载项。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 1ec844af555ab55d7592a9d764c6173ea9780016
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938422"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>在 Office 网页版中调试加载项

您可以在并非运行 Windows 或 Office 2013 或 Office 2016 桌面客户端的计算机上构建和调试外接程序，例如，如果您正在使用 Mac 进行开发。本文介绍如何使用 Office Online 测试和调试您的外接程序。 本文介绍了如何使用 Office 网页版来测试和调试加载项。 

## <a name="prerequisites"></a>先决条件

首先，请执行以下操作：

- 如果Microsoft 365或没有对网站的访问权限，请获取一个SharePoint帐户。

  > [!NOTE]
  > 若要获取为期 90 天的免费可续订开发人员Microsoft 365订阅，请加入我们的Microsoft 365[计划](https://developer.microsoft.com/office/dev-program)。 有关如何[Microsoft 365](/office/developer-program/office-365-developer-program) Microsoft 365开发人员计划并配置订阅的分步说明，请参阅 Microsoft 365 开发人员计划文档。

- 在 SharePoint Online 上SharePoint目录。 应用程序目录是 SharePoint Online 中的专用网站集，Office外接程序的文档库。如果你有自己的网站SharePoint，可以设置应用程序目录文档库。 有关详细信息，请参阅上的将任务窗格和内容加载项发布到[应用程序SharePoint。](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a>在 Excel 网页版或 Word 网页版中调试加载项

若要使用 Office 网页版调试加载项，请执行以下操作：

1. 将加载项部署到支持 SSL 的服务器上。

    > [!NOTE]
    > 建议使用 [Yeoman 生成器](https://github.com/OfficeDev/generator-office)创建和托管加载项。

2. 在 [加载项清单文件](../develop/add-in-manifests.md)中，将 **SourceLocation** 元素值更新为包括绝对 URI，而不是相对 URI。例如：

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. 将清单上传到 SharePoint 上应用程序目录中的 Office 加载项文档库。

4. 从Excel启动Word web 版启动器启动Microsoft 365或打开一个新文档。

5. 在“插入”选项卡上选择“我的外接程序”或“Office 外接程序”以插入您的外接程序并在应用程序中进行测试。

6. 使用常用浏览器工具调试器调试加载项。

## <a name="potential-issues"></a>潜在问题

以下是在调试时可能会遇到的一些问题。

- 你看到的一些 JavaScript 错误可能源自 Office 网页版。

- 浏览器可能会显示无效证书错误，你需要忽略此错误。 执行此操作的过程因浏览器而异，而且用于执行此操作的各种浏览器的 UI 会定期进行更改。 有关说明，可搜索浏览器的“帮助”或“联机搜索”。 （例如，搜索“Microsoft Edge 无效证书警告”。）大多数浏览器在“警告”页面上都有一个链接，可以通过此链接单击进入“加载项”页。 例如，Microsoft Edge 有一个链接“转到网页（不推荐）”。 但是每次加载项重新加载时，通常都必须通过此链接来完成。 如需更长久地忽略，请参阅建议的帮助。

- 如果你在代码中设置了断点，Office 网页版可能会抛出错误，指明它无法保存。

## <a name="see-also"></a>另请参阅

- [Office 加载项开发最佳做法](../concepts/add-in-development-best-practices.md)
- [AppSource 验证策略](/legal/marketplace/certification-policies)  
- [创建有效的 AppSource 应用和加载项](/office/dev/store/create-effective-office-store-listings)  
- [排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)

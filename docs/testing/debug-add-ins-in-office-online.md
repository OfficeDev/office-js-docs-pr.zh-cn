---
title: 在 Office 网页版中调试加载项
description: 如何使用 Office 网页版来测试和调试加载项。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: f7ef3fa3d6389629e28b428b9bdbe3b128896b1f
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094489"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>在 Office 网页版中调试加载项

您可以在并非运行 Windows 或 Office 2013 或 Office 2016 桌面客户端的计算机上构建和调试外接程序，例如，如果您正在使用 Mac 进行开发。本文介绍如何使用 Office Online 测试和调试您的外接程序。 本文介绍了如何使用 Office 网页版来测试和调试加载项。 

## <a name="prerequisites"></a>先决条件

首先，请执行以下操作：

- 获取 Microsoft 365 开发人员帐户（如果还没有）或有权访问 SharePoint 网站。

  > [!NOTE]
  > To get a free, 90-day renewable Microsoft 365 developer subscription, join our [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program). See the [Microsoft 365 developer program documentation](/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Microsoft 365 developer program and configure your subscription.

- Set up an app catalog on SharePoint Online. An app catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an app catalog document library. For more information, see [Publish task pane and content add-ins to an app catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a>在 Excel 网页版或 Word 网页版中调试加载项

若要使用 Office 网页版调试加载项，请执行以下操作：

1. 将加载项部署到支持 SSL 的服务器上。

    > [!NOTE]
    > 建议使用 [Yeoman 生成器](https://github.com/OfficeDev/generator-office)创建和托管加载项。

2. In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. 将清单上传到 SharePoint 上应用程序目录中的 Office 加载项文档库。

4. 从 Microsoft 365 中的应用启动器启动 Excel 或 Word，然后打开一个新文档。

5. 在“插入”选项卡上选择“我的外接程序”**** 或“Office 外接程序”**** 以插入您的外接程序并在应用程序中进行测试。

6. 使用常用浏览器工具调试器调试加载项。

## <a name="potential-issues"></a>潜在问题

下面介绍了一些在调试过程中可能会遇到的问题：

- 你看到的一些 JavaScript 错误可能源自 Office 网页版。

- 浏览器可能会显示无效证书错误，你需要忽略此错误。 执行此操作的过程因浏览器而异，而且用于执行此操作的各种浏览器的 UI 会定期进行更改。 有关说明，可搜索浏览器的“帮助”或“联机搜索”。 （例如，搜索“Microsoft Edge 无效证书警告”。）大多数浏览器在“警告”页面上都有一个链接，可以通过此链接单击进入“加载项”页。 例如，Microsoft Edge 有一个链接“转到网页（不推荐）”。 但是每次加载项重新加载时，通常都必须通过此链接来完成。 如需更长久地忽略，请参阅建议的帮助。

- 如果你在代码中设置了断点，Office 网页版可能会抛出错误，指明它无法保存。

## <a name="see-also"></a>另请参阅

- [Office 加载项开发最佳做法](../concepts/add-in-development-best-practices.md)
- [AppSource 验证策略](/legal/marketplace/certification-policies)  
- [创建有效的 AppSource 应用和加载项](/office/dev/store/create-effective-office-store-listings)  
- [排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)

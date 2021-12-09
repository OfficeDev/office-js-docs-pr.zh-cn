---
title: 部署和发布 Office 加载项
description: 部署 Office 加载项以进行测试或分发给用户的方法和选项。
ms.date: 12/07/2021
ms.localizationpriority: high
ms.openlocfilehash: 81c02a36becb9ef3244f7754dda44d064cdd9925
ms.sourcegitcommit: e392e7f78c9914d15c4c2538c00f115ee3d38a26
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/08/2021
ms.locfileid: "61331076"
---
# <a name="deploy-and-publish-office-add-ins"></a>部署和发布 Office 加载项

可以使用几种方法之一来部署 Office 外接程序，以用于对用户进行测试或分发：

|**方法**|**Use...**|
|:---------|:------------|
|[旁加载](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)|在开发过程中，测试 Windows、iPad、Mac 或浏览器中运行的加载项。（不适用于生产加载项。）|
|[网络共享](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|作为开发过程的一部分，用于在将加载项发布到本地主机以外的服务器后，测试在 Windows 上运行的加载项。 （不适用于生产版加载项，也不适用于在 iPad、Mac 或 Web 上进行测试。）|
|[AppSource](/office/dev/store/submit-to-appsource-via-partner-center)|用于向用户公开分发加载项。|
|[ Microsoft 365 管理中心](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)|在云部署中，使用 Microsoft 365 管理中心将加载项分发给组织中的用户。 这是通过[集成应用](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)或[集中部署](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)完成的。 |
|[SharePoint 目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|在本地环境中，用于向组织用户分发加载项。|
|[Exchange 服务器](#outlook-add-in-deployment)|在本地或在线环境中，用于向用户分发 Outlook 加载项。|

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="deployment-options-by-office-application-and-add-in-type"></a>按 Office 应用程序和加载项类型划分的部署选项

可用的部署选项具体取决于你面向的 Office 应用程序以及所创建的加载项的类型。

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Word、Excel 和 PowerPoint 加载项的部署选项

| 扩展点 | 旁加载 | 网络共享 | AppSource | Microsoft 365 管理中心 | SharePoint 目录\* |
|:----------------|:-----------:|:-------------:|:---------:|:--------------------------:|:--------------------:|
| 内容         | X           | X             | X         | X                          | X                    |
| 任务窗格       | X           | X             | X         | X                          | X                    |
| 命令         | X           | X             | X         | X                          |                      |

&#42; SharePoint 目录不支持 Mac 版 Office。

### <a name="deployment-options-for-outlook-add-ins"></a>Outlook 加载项的部署选项

| 扩展点 | 旁加载 | AppSource | Exchange 服务器 |
|:----------------|:-----------:|:---------:|:---------------:|
| 邮件应用        | X           | X         | X               |
| 命令         | X           | X         | X               |

## <a name="production-deployment-methods"></a>产品部署方法

以下各部分提供了有关向组织中的用户分发生产版 Office 加载项的最常用部署方法的其他信息。

有关最终用户如何获取、插入和运行加载项的信息，请参阅[开始使用 Office 加载项](https://support.microsoft.com/office/82e665c4-6700-4b56-a3f3-ef5441996862)。

### <a name="integrated-apps-via-the-microsoft-365-admin-center"></a>通过 Microsoft 365 管理中心集中应用

通过 Microsoft 365 管理中心，管理员可以为组织中的用户和组轻松部署 Office 加载项。 通过管理中心部署加载项后，用户可立即在其 Office 应用程序中使用此加载项，而无需进行客户端配置。 可以使用集中应用部署内部加载项，以及 ISV 提供的加载项。 集成应用还显示管理员加载项，和由同一 ISV 捆绑在一起的其他应用，以便能够充分体验 Microsoft 365 平台。

将 Office 加载项、Teams 应用、SPFx 应用和[其他应用](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps#what-apps-can-i-deploy-from-integrated-apps)链接在一起时，可为客户创建单个软件即服务 （SaaS） 产品/服务。 有关此流程的通用信息，请参阅[如何为商业市场计划 SaaS 产品/服务](/azure/marketplace/plan-saas-offer)。 有关如何创建集成应用的详细信息，请参阅[配置 Microsoft 365 应用集成](/azure/marketplace/create-new-saas-offer#configure-microsoft-365-app-integration)。

有关集成应用部署流程的详细信息，请参阅[集成应用门户中的合作伙伴测试和部署Microsoft 365 应用](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)。

> [!IMPORTANT]
> 主权云或政府云中的客户无权访问集成应用。 它们将改用集中部署。 集中部署是类似的部署方法，但不会向管理员公开连接的加载项和应用。有关详细信息，请参阅 [确定加载项的集中部署是否适用于组织](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)。

### <a name="sharepoint-app-catalog-deployment"></a>SharePoint 应用目录部署

SharePoint 应用目录是特殊网站集，创建后可用于托管 Word、Excel 和 PowerPoint 加载项。由于 SharePoint 目录不支持在清单的 `VersionOverrides` 节点中实现的新加载项功能（包括加载项命令），因此建议尽可能通过管理中心进行集中部署。通过 SharePoint 目录部署的加载项命令默认在任务窗格中打开。

如果要在本地环境中部署外接程序，请使用 SharePoint 目录。有关详细信息，请参阅[将任务窗格和内容外接程序发布到 SharePoint 目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。

> [!NOTE]
> SharePoint 目录不支持 Mac 版 Office。 若要向 Mac 客户端部署 Office 加载项，必须将其提交到 [AppSource](/office/dev/store/submit-to-the-office-store)。

### <a name="outlook-add-in-deployment"></a>Outlook 加载项部署

对于不使用 Azure AD 标识服务的本地和联机环境，可以通过 Exchange 服务器部署 Outlook 外接程序。

Outlook 外接程序部署需要以下内容：

- Microsoft 365、Exchange Online 或 Exchange Server 2013 或更高版本
- Outlook 2013 或更高版本

要将加载项分配给租户，请使用 Exchange 管理中心从文件或 URL 直接上传清单，或从 AppSource 添加加载项。 若要将加载项分配给单个用户，必须使用 Exchange PowerShell。 有关详细信息，请参阅 [Exchange Server 中 Outlook 的加载项](/exchange/add-ins-for-outlook-2013-help)。

## <a name="see-also"></a>另请参阅

- [旁加载 Outlook 加载项以供测试](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [提交到 AppSource][AppSource]
- [Office 加载项的设计准则](../design/add-in-design.md)
- [创建有效的 AppSource 一览](/office/dev/store/create-effective-office-store-listings)
- [排查 Office 加载项中的用户错误](../testing/testing-and-troubleshooting.md)
- [什么是 Microsoft 商业市场？](/azure/marketplace/overview)

[AppSource]: /office/dev/store/submit-to-appsource-via-partner-center

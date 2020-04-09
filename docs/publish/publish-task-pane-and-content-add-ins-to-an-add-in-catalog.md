---
title: 将任务窗格和内容加载项发布到 SharePoint 应用程序目录
description: 为使组织内的用户可访问 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的应用程序目录中。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 5557dd31e829fac2c2dbd421200da46a5c3b9b99
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185587"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a>将任务窗格和内容加载项发布到 SharePoint 应用程序目录

应用程序目录是 SharePoint Web 应用程序或 SharePoint Online 租户中的专用网站集，用于托管 Office 和 SharePoint 加载项的文档库。若要向组织用户分发 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的应用程序目录。如果管理员将应用程序目录注册为受信任的目录，用户就可以通过 Office 客户端应用程序中的插入 UI 插入加载项。

> [!IMPORTANT]
> - SharePoint 上的应用程序目录不支持在[加载项清单](../develop/add-in-manifests.md)的 `VersionOverrides` 节点中实现的加载项功能（如加载项命令）。
> - 如果面向的是云或混合环境，建议通过 [Office 365 管理中心使用集中部署](../publish/centralized-deployment.md)来发布加载项。
> - Mac 版 Office 不支持 SharePoint 上的应用程序目录。 若要向 Mac 客户端部署 Office 加载项，必须将其提交到 [AppSource](/office/dev/store/submit-to-the-office-store)。

## <a name="create-an-app-catalog"></a>创建应用程序目录

完成以下某个部分中的步骤，以使用本地 SharePoint Server 或 Office 365 创建应用程序目录。

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a>为本地 SharePoint Server 创建应用程序目录

若要创建 SharePoint 应用程序目录，请按照[配置 Web 应用程序的应用程序目录网站](/sharepoint/administration/manage-the-app-catalog)中的说明进行操作。

创建应用程序目录后，请按照相关步骤[发布 Office 加载项](#publish-an-office-add-in)。

### <a name="to-create-an-app-catalog-on-office-365"></a>在 Office 365 上创建应用程序目录

若要创建 SharePoint 应用程序目录，请按照[create The App catalog site collection](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection)中的说明操作。 创建应用程序目录后，请按照下一节中的步骤发布 Office 外接程序。

## <a name="publish-an-office-add-in"></a>发布 Office 加载项

完成以下某个部分中的步骤，以将 Office 加载项发布到 Office 365 或本地 SharePoint Server 上的应用程序目录。

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a>将 Office 加载项发布到 Office 365 上的 SharePoint 应用程序目录

1. 转到[新的 SharePoint 管理中心的“活动站点”页面](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true)，然后使用在组织中具有[管理员权限](/sharepoint/sharepoint-admin-role)的帐户进行登录。

>[!NOTE]
>如果使用的是 Office 365 Germany，请[登录 Microsoft 365 管理中心](https://go.microsoft.com/fwlink/p/?linkid=848041)，然后浏览到 SharePoint 管理中心并打开“更多功能”页面。 <br>如果使用的是由世纪互联（中国）运营的 Office 365，请[登录 Microsoft 365 管理中心](https://go.microsoft.com/fwlink/p/?linkid=850627)，然后浏览到 SharePoint 管理中心并打开“更多功能”页面。
 
2. 通过在 "URL" 列中选择应用程序目录网站的 URL 来打开该网站。 

>[!NOTE]
>如果你刚刚在上一节中创建了应用程序目录网站，可能需要几分钟时间才能完成网站设置。

3. 选择“**分发 Office 应用程序**”。
4. 在“**Office 应用程序**”页中，选择“**新建**”。
5. 在“**添加文档**”对话框中，选择“**选择文件**”按钮。
6. 找到并指定要上传的“[清单文件](../develop/add-in-manifests.md)”，并选择“**打开**”。
7. 在“**添加文档**”对话框中，选择“**确定**”。

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a>使用本地 SharePoint Server 将加载项发布到应用程序目录

1. 打开“**管理中心**”页。
2. 在左侧的任务窗格中，选择“**应用程序**”。
3. 在“**应用程序**”页的“**应用程序管理**”下方，选择“**管理应用程序目录**”。
4. 在“**管理应用程序目录**”页上，确保在“**Web 应用程序**”选择器中选择了正确的 Web 应用程序。
5. 选择“**网站 URL**”下的 URL 以打开应用程序目录网站。
6. 选择“**分发 Office 应用程序**”。
7. 在“**Office 应用程序**”页中，选择“**新建**”。
8. 在“**添加文档**”对话框中，选择“**选择文件**”按钮。
9. 找到并指定要上传的“[清单文件](../develop/add-in-manifests.md)”，并选择“**打开**”。
10. 在“**添加文档**”对话框中，选择“**确定**”。

## <a name="insert-office-add-ins-from-the-app-catalog"></a>从应用程序目录插入 Office 加载项

对于联机 Office 应用程序，你可以通过完成以下步骤从应用程序目录中找到 Office 加载项。

1. 打开联机 Office 应用程序（Excel、PowerPoint 或 Word）。
2. 创建或打开文档。
3. 选择“**插入**” > “**加载项**”。
4. 在“Office 加载项”对话框中，选择“**我的组织**”选项卡。此时将列出 Office 加载项。
5. 选择 Office 加载项，然后选择“**添加**”。

对于桌面上的 Office 应用程序，你可以通过完成以下步骤从应用程序目录中找到 Office 加载项。

1. 打开桌面版 Office 应用程序（Excel、Word 或 PowerPoint）
2. 选择“**文件**” > “**选项**” > “**信任中心**” > “**信任中心设置**” > “**受信任的加载项目录**”。
3. 在“**目录 URL**”框中输入 SharePoint 应用程序目录的 URL，然后选择“**添加目录**”。
    使用较短形式的 URL。 例如，如果 SharePoint 应用程序目录的 URL 为：
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`
    
    仅指定父网站集的 URL：
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
4. 关闭并重新打开 Office 应用程序。 
5. 选择“**插入**” > “**获取加载项**”。
4. 在“Office 加载项”对话框中，选择“**我的组织**”选项卡。此时将列出 Office 加载项。
5. 选择 Office 加载项，然后选择“**添加**”。

或者，管理员可以使用组策略在 SharePoint 上指定应用目录。 [适用于 Office 365 专业增强版、Office 2019 和 Office 2016 的管理模板文件 (ADMX/ADML)](https://www.microsoft.com/download/details.aspx?id=49030) 中提供了相关的策略设置，可在 **User Configuration\Policies\Administrative Templates\Microsoft Office 2016\Security Settings\Trust Center\Trusted Catalogs** 下找到。

---
title: 将任务窗格和内容加载项发布到 SharePoint 应用程序目录
description: 为使组织内的用户可访问 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的应用程序目录中。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: e9a600cd807379e9c55f2fc98bb4f2d71552058f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325303"
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

1. 转到 Microsoft 365 管理中心。 有关如何查找管理中心的信息，请参阅[关于 Microsoft 365 管理中心](/office365/admin/admin-overview/about-the-admin-center)。

2. 在 Microsoft 365 管理中心页面上，展开“**管理中心**”列表，然后选择“**SharePoint**”。

    > [!NOTE]
    > 需要使用经典 SharePoint 管理中心才能创建目录。 如果位于新的 SharePoint 管理中心，请在左侧窗格中选择“**经典 SharePoint 管理中心**”。

3. 在左侧任务窗格中，选择 "**应用程序**"。

4. 在“**应用程序**”页面上，选择“**应用程序目录**”。
    > [!NOTE]
    > 如果已创建应用程序目录并且它显示在此页面上，则你可以跳过其余步骤并转至本文下一章节，将你的加载项发布到目录。

5. 在“**应用程序目录网站**”页上，选择“**确定**”以接受默认选项并创建新的应用程序目录网站。

6. 在“**创建应用程序目录网站集**”页上，指定应用程序目录网站的标题。

7. 指定**网站地址**。

8. 指定**管理员**。

9. 将**服务器资源配额**设为 0（零）。 （服务器资源配额与限制性能不佳的沙盒解决方案有关，但你不会在应用程序目录网站上安装任何沙盒解决方案。）

10. 选择“**确定**”。

## <a name="publish-an-office-add-in"></a>发布 Office 加载项

完成以下某个部分中的步骤，以将 Office 加载项发布到 Office 365 或本地 SharePoint Server 上的应用程序目录。

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a>将 Office 加载项发布到 Office 365 上的 SharePoint 应用程序目录

1. 转到 Microsoft 365 管理中心。 有关如何查找管理中心的信息，请参阅[关于 Microsoft 365 管理中心](/office365/admin/admin-overview/about-the-admin-center)。
2. 在 Microsoft 365 管理中心页面上，展开“**管理中心**”列表，然后选择“**SharePoint**”。
    > [!NOTE]
    > 需要使用经典 SharePoint 管理中心才能创建目录。 如果位于新的 SharePoint 管理中心，请在左侧窗格中选择“**经典 SharePoint 管理中心**”。
3. 在左侧任务窗格中，选择 "**应用程序**"。
4. 在“**应用程序**”页面上，选择“**应用程序目录**”。
5. 选择“**分发 Office 应用程序**”。
6. 在“**Office 应用程序**”页中，选择“**新建**”。
7. 在“**添加文档**”对话框中，选择“**选择文件**”按钮。
8. 找到并指定要上传的“[清单文件](../develop/add-in-manifests.md)”，并选择“**打开**”。
9. 在“**添加文档**”对话框中，选择“**确定**”。

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

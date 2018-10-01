---
title: 将任务窗格和内容加载项发布到 SharePoint 目录
description: 若要使其组织内的用户可以访问 Office 加载项，管理员可以将 Office 加载项的清单文件上载到其组织的加载项目录。
ms.date: 01/23/2018
ms.openlocfilehash: 5ba6a54c4540f79c65082cd7de3b76f300831341
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348119"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a>将任务窗格和内容加载项发布到 SharePoint 目录

加载项目录是 SharePoint Web 应用或 SharePoint Online 租赁中的专用网站集，用于托管 Office 和 SharePoint 加载项的文档库。若要向组织用户分发 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的加载项目录。如果管理员将加载项目录注册为受信任的目录，用户就可以通过 Office 客户端应用中的插入 UI 插入加载项。

> [!IMPORTANT]
> - SharePoint 上的加载项目录不支持在[加载项清单](../develop/add-in-manifests.md)的 `VersionOverrides` 节点中实现的加载项功能（如加载项命令）。
> - 如果面向的是云或混合环境，建议通过 [Office 365 管理中心使用集中部署](../publish/centralized-deployment.md)来发布加载项。
> - Office for Mac 不支持 SharePoint 目录。 若要向 Mac 客户端部署 Office 加载项，您必须将它们提交给 [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)。   

## <a name="set-up-an-add-in-catalog"></a>设置加载项目录

完成以下部分之一中的步骤，以在 SharePoint 或 Office 365 上设置加载项目录。

### <a name="to-set-up-an-add-in-catalog-for-on-premises-sharepoint"></a>在本地 SharePoint 上设置加载项目录

> [!NOTE]
> 本地 SharePoint 中的 UI 仍将加载项称为**应用程序**。

1. 浏览到**管理中心站点**。
    
2. 在左侧的任务窗格中，选择**应用程序**。
    
3. 在**应用**页面的**应用管理**下，选择**管理应用目录**。
    
4. 在**管理应用程序目录**页上，确保在**Web 应用程序选择器**中选择了正确的 Web 应用程序。
    
5. 选择**查看站点设置**。
    
6. 在**网站设置**页上选择**网站集管理员**以指定网站集管理员，然后选择**确定**。
    
7. 要向用户授予站点权限，请选择**站点权限**，然后选择**授予权限**。
    
8. 在**共享‘应用程序目录网站’** 对话框中，指定一个或多个网站用户，为他们设置相应的权限，选择性地设置其他选项，然后选择**共享**。
    
9. 要向 Office 加载项的加载项目录添加加载项，请选择 **Office 应用程序**。

### <a name="to-set-up-an-add-in-catalog-on-office-365"></a>在 Office 365 上设置加载项目录

1. 在 Office 365 管理中心页上，选择**管理员**，然后选择 **SharePoint**。
    
2. 在左侧的任务窗格中，选择**加载项**。
    
3. 在**加载项**页上，选择**加载项目录**。
    
4. 在**加载项目录网站**页上，选择**确定**以接受默认选项，并新建加载项目录网站。
    
5. 在**创建加载项目录网站集**页上，指定加载项目录站点的标题。
    
6. 指定网站地址。
    
7. 将**存储配额**设置为可能的最低值（当前为 110）。你将仅在该网站集上安装加载项包，它们非常小。
    
8. 将**服务器资源配额**设置为 0（零）。（服务器资源配额与限制性能不佳的沙盒解决方案有关，但你不会在加载项目录网站上安装任何沙盒解决方案。）
    
9. 选择**确定**。
    
10. 若要将加载项添加到加载项目录网站，请转到刚刚创建的网站。在左侧导航窗格中，依次选择 **Office 加载项**和**新加载项**，以上传 Office 加载项清单文件。

## <a name="publish-an-add-in-to-an-add-in-catalog"></a>将加载项发布到加载项目录

若要将加载项发布到加载项目录，请完成以下步骤。

1. 浏览到加载项目录：

    - 打开 SharePoint 管理中心主页。
    
    - 选择**加载项**。
    
    - 选择**管理加载项目录**。
    
    - 依次选择所提供的链接和左侧导航栏上的 **Office 加载项**。
    
2. 选择**单击添加新项**链接。
    
3. 选择**浏览**，再指定要上传的[清单](../develop/add-in-manifests.md)。
    
    此目录中的内容和任务窗格加载项现在可从**Office 加载项**对话框提供。若要访问这些加载项，请在**插入**选项卡上选择**我的加载项**，然后选择**我的组织**。

## <a name="end-user-experience-with-the-add-in-catalog"></a>加载项目录的最终用户体验

最终用户可以通过完成以下步骤来访问 Office 应用程序中的加载项目录：

1. 在 Office 应用程序中，转到**文件** > **选项** > **信任中心** > **信任中心设置** > **受信任的加载项目录**。
    
2. 指定加载项目录的_父级 SharePoint 网站集_的 URL。 
    
    例如，如果“Office 加载项”目录的 URL 是：
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    仅指定父网站集的 URL：
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. 关闭并重新打开 Office 应用程序。加载项目录将在 **Office 加载项**对话框中提供。

或者，管理员可以使用组策略在 SharePoint 上指定 Office 加载项目录。 有关详细信息，请参阅[使用组策略管理用户可以安装和使用 Office 加载项的方式](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office)一节。

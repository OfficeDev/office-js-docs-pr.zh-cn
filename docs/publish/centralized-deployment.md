---
title: 通过 Microsoft 365 管理中心使用集中部署发布 Office 外接程序
description: 了解如何使用集中部署来部署内部加载项以及 Isv 提供的外接程序。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 0e99742be87b477b7c78295d08539de924f02466
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094251"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-microsoft-365-admin-center"></a>通过 Microsoft 365 管理中心使用集中部署发布 Office 外接程序

Microsoft 365 管理中心使管理员可以轻松地将 Office 加载项部署到其组织内的用户和组。 在管理员通过管理中心部署加载项后，用户可以立即在 Office 应用中使用加载项，而无需进行任何客户端配置。 通过集中部署，可以部署内部加载项和 ISV 提供的加载项。

Microsoft 365 管理中心目前支持以下方案。

- 为个人、组或组织集中部署新的和更新的加载项。
- 部署到多个客户端平台，包括 Windows、Mac 和 web。 对于 Outlook，也支持部署到 iOS 和 Android。 但是，在支持用户在 iPad 上安装 Excel、Outlook、Word 和 PowerPoint 外接程序时 (，**不**支持集中部署到 ipad。 ) 
- 到英语语言租户和全球范围租户的部署。
- 部署云托管的加载项。
- 部署托管在防火墙内的加载项。
- 部署 AppSource 加载项。
- 当用户启动 Office 应用时自动为用户安装加载项。
- 当管理员禁用或删除加载项，或者将用户从 Azure Active Directory 或从已部署加载项的组中删除时，则自动为用户删除该加载项。

如果组织满足使用集中部署的所有要求，则建议采用集中部署，以便 Microsoft 365 管理员在组织内部署 Office 加载项。 有关如何确定组织是否可以使用集中部署的信息，请参阅[确定加载项的集中部署是否适用于你的 Microsoft 365 组织](/office365/admin/manage/centralized-deployment-of-add-ins)。

> [!NOTE]
> 在不连接到 Microsoft 365 的本地环境中，或者若要部署 SharePoint 外接程序或面向 Office 2013 的 Office 外接程序，请使用[SharePoint 应用程序目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。 若要部署 COM/VSTO 加载项，请使用 ClickOnce 或 Windows Installer，如[部署 Office 解决方案](/visualstudio/vsto/deploying-an-office-solution)中所述。

## <a name="recommended-approach-for-deploying-office-add-ins"></a>部署 Office 加载项的推荐方法

Consider deploying Office Add-ins in a phased approach to help ensure that the deployment goes smoothly. We recommend the following plan:

1. Deploy the add-in to a small set of business stakeholders and members of the IT department. If the deployment is successful, move on to step 2.

2. Deploy the add-in to a larger set of individuals within the business who will be using the add-in. If the deployment is successful, move on to step 3.

3. 为所有将使用加载项的个人部署加载项。

根据目标受众的规模，可能需要在此过程中添加步骤或删除步骤。

## <a name="publish-an-office-add-in-via-centralized-deployment"></a>通过集中部署发布 Office 加载项

在开始之前，请确认您的组织满足使用集中部署的所有要求，如[确定外接程序的集中部署是否适用于您的 Microsoft 365 组织](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)中所述。

如果组织满足所有要求，请完成以下步骤以通过集中部署发布 Office 加载项：

1. 使用你的工作或教育帐户登录 Microsoft 365。
2. 选择左上角的应用启动器图标，然后选择“**管理员**”。
3. 在导航菜单中，按“**显示更多内容**”，然后选择“**设置**” > “**服务和加载项**”。
4. 如果在宣布新的 Microsoft 365 管理中心的页面顶部看到一条消息，请选择该消息以转到管理中心预览 (请参阅[关于 Microsoft 365 管理中心](/microsoft-365/admin/admin-overview/about-the-admin-center)) 。
5. 在页面顶部选择“**部署加载项**”。
6. 查看要求后，请选择“**下一步**”。
7. 在“**集中部署**”页面上，选择以下选项之一：

    - **我想从 Office 应用商店添加加载项。**
    - **我在此设备上具有清单文件 (.xml)。** 对于此选项，请选择“浏览”**** 以找到想要使用的清单文件 (.xml)。
    - **I have a URL for the manifest file.** For this option, type the manifest's URL in the field provided.

    ![Microsoft 365 管理中心中的新加载项对话框](../images/new-add-in.png)

8. 如果选择了此选项以从 Office 应用商店添加某个加载项，请选择该加载项。 可以通过“**为你推荐**”、“**评级**”或“**名称**”类别，查看可用的加载项。 仅能从 Office 应用商店添加免费加载项。 目前不支持添加付费加载项。

    > [!NOTE]
    > 使用 Office 应用商店选项，无需干预，用户即可自动获得加载项的更新和增强功能。

    ![在 Microsoft 365 管理中心中选择外接程序对话框](../images/select-an-add-in.png)

9. 查看加载项详细信息、隐私策略和许可条款后，选择 "**继续**"。

    ![Microsoft 365 管理中心中的所选加载项页面](../images/selected-add-in-admin-center.png)

10. 在 "**分配用户**" 页上，选择 "**所有人**"、"**特定用户/组**" 或 "**仅我自己**"。 使用“搜索”框查找要向其部署加载项的用户和组。 对于 Outlook 外接程序，您还可以选择 "已**修复**"、"**可用**" 或 "**可选**" 部署方法。

    ![在 Microsoft 365 管理中心管理有权访问和部署方法的成员](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > 用于加载项的[单一登录 (SSO)](../develop/sso-in-office-add-ins.md) 系统目前处于预览状态，不应用于生产加载项。部署使用 SSO 的加载项时，分配的用户和组也将与共享相同 Azure App ID 的加载项共享。 对用户分配进行的任何更改也会应用于这些加载项。相关加载项显示在此页面上。 仅对于 SSO 加载项，此页面将显示加载项所需的 Microsoft Graph 权限的列表。

11. 完成后，选择 "**部署**"。 此过程可能最多用时 3 分钟。 然后，按“**下一步**”完成演练。 现在，可以看到此加载项与其他应用一起显示在 Office 365 中。

    > [!NOTE]
    > 当管理员选择 "**部署**" 时，将为所有用户授予 "同意"。

    ![Microsoft 365 管理中心中的应用程序列表](../images/citations.png)

> [!TIP]
> 为组织中的用户和/或组部署新加载项时，请考虑向他们发送一封电子邮件，说明加载项的应用场景和使用方式，并添加相关帮助内容、FAQ 或其他支持资源的链接。

## <a name="considerations-when-granting-access-to-an-add-in"></a>授予加载项的访问权限时的注意事项

Admins can assign an add-in to everyone in the organization or to specific users and/or groups within the organization. The following list describes the implications of each option:

- **Everyone**: As the name implies, this option assigns the add-in to every user in the tenant. Use this option sparingly and only for add-ins that are truly universal to your organization.

- **Users**: If you assign an add-in to individual users, you'll need to update the Central Deployment settings for the add-in each time you want to assign it additional users. Likewise, you'll need to update the Central Deployment settings for the add-in each time you want to remove a user's access to the add-in.

- **组**：如果将加载项分配给组，则会自动为被添加到此组的用户分配此加载项。 同样，当将某个用户从组中删除时，此用户将自动失去对此加载项的访问权限。 无论在哪种情况下，Microsoft 365 管理员都不需要执行任何其他操作。

In general, for ease of maintenance, we recommend assigning add-ins by using groups whenever possible. However, in situations where you want to restrict add-in access to a very small number of users, it may be more practical to assign the add-in to specific users.

## <a name="add-in-states"></a>加载项状态

下表介绍了加载项的不同状态。

|状态|此状态如何出现|影响|
|-----|--------------------|------|
|**活动**|管理员已上传加载项并已将其分配给用户和/或组。|已为其分配加载项的用户和/或组，可在相关的 Office 客户端中找到它。|
|**已禁用**|管理员已禁用加载项。|Users and/or groups assigned to the add-in no longer have access to it. If the add-in state is changed from **Turned off** to **Active**, the users and groups will regain access to it.|
|**已删除**|管理员已删除加载项。|已为其分配加载项的用户和/或组不再能够访问它。|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a>更新通过集中部署发布的 Office 加载项

After an Office Add-in has been published via Centralized Deployment, any changes made to the add-in's web application will automatically be available to all users as soon as those changes are implemented in the web application. Changes made to an add-in's [XML manifest file](../develop/add-in-manifests.md), for example, to update the add-in's icon, text, or add-in commands, happen as follows:

- **业务线外接程序**：如果管理员在通过 Microsoft 365 管理中心实施集中化部署时显式上载了清单文件，则管理员必须上载包含所需更改的新清单文件。 上传更新后的清单文件后，加载项就会在下次相关 Office 应用启动时更新。

  > [!NOTE]
  > 管理员无需删除 LOB 加载项即可进行更新。 在 "外接程序" 部分中，管理员只需选择 LOB 外接程序，然后按右下角的 "**更新外接程序"** 按钮，即可调用此功能。
  > 
  > ![屏幕截图显示了 Microsoft 365 管理中心中的更新外接端对话框](../images/update-add-in-admin-center.png)

- **Office 应用商店外接程序**：如果管理员通过 Microsoft 365 管理中心实施集中部署时从 Office 应用商店中选择了外接程序，并且 Office 应用商店中的外接程序更新，则外接程序将在以后通过集中部署进行更新。 加载项会在下次相关 Office 应用启动时更新。

## <a name="end-user-experience-with-add-ins"></a>加载项最终用户体验

通过集中部署发布加载项后，最终用户可以在加载项支持的任何平台上开始使用它。

If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. In the following example, the command **Search Citation** appears in the ribbon for the **Citations** add-in.

![屏幕截图显示在引文加载项中突出显示 "搜索引文" 命令的 Office 应用程序功能区部分](../images/search-citation.png)

如果加载项不支持加载项命令，用户可以通过执行以下操作将其添加到 Office 应用程序中：

1. 在 Word 2016 或更高版本、Excel 2016 或更高版本，或 PowerPoint 2016 或更高版本，选择“**插入**” > “**我的加载项**”。
2. 在加载项窗口中选择“**管理托管**”选项卡。
3. 选择加载项，然后选择“添加”****。

    ![Screenshot shows the Admin Managed tab of the Office Add-ins page of an Office application. The Citations add-in is shown on the tab.](../images/office-add-ins-admin-managed.png)

但是，对于 Outlook 2016 或更高版本，用户可以执行以下操作：

1. 在 Outlook 中，选择“**开始**” > “**应用商店**”。
2. 选择“加载项”选项卡下的“**管理员管理**”选项卡。
3. 选择加载项，然后选择“**添加**”。

    ![屏幕截图显示了 Outlook 应用程序的“应用商店”页面的管理员管理区域。](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a>另请参阅

- [确定加载项的集中部署是否适用于你的 Microsoft 365 组织](/office365/admin/manage/centralized-deployment-of-add-ins)

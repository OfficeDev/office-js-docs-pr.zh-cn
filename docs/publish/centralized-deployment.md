---
title: 通过 Microsoft 365 管理中心使用集中部署发布 Office 外接程序
description: 了解如何使用集中部署来部署内部外接程序以及 ISV 提供的外接程序。
ms.date: 01/12/2021
localization_priority: Normal
ms.openlocfilehash: c1f36d1ad640adbecdd3338200e742e76831a67a
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839752"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-microsoft-365-admin-center"></a>通过 Microsoft 365 管理中心使用集中部署发布 Office 外接程序

Microsoft 365 管理中心使管理员能够轻松地将 Office 加载项部署到其组织内的用户和组。 通过管理中心部署加载项后，用户可立即在其 Office 应用程序中使用此加载项，而无需进行客户端配置。 可以通过集中部署来部署内部加载项以及 ISV 提供的加载项。

Microsoft 365 管理中心当前支持以下方案。

- 为个人、组或组织集中部署新的和更新的加载项。
- 部署到多个客户端平台，包括 Windows、Mac 和 Web。 对于 Outlook，也支持部署到 iOS 和 Android。  (，尽管支持在 iPad 上安装 Excel、Outlook、Word 和 PowerPoint 加载项，但不支持集中部署到 iPad。) 
- 到英语语言租户和全球范围租户的部署。
- 部署云托管的加载项。
- 部署托管在防火墙内的加载项。
- 部署 AppSource 加载项。
- 当用户启动 Office 应用时自动为用户安装加载项。
- 当管理员禁用或删除加载项，或者将用户从 Azure Active Directory 或从已部署加载项的组中删除时，则自动为用户删除该加载项。

建议 Microsoft 365 管理员使用集中部署在组织中部署 Office 加载项，但组织必须满足使用集中部署的所有要求。 若要了解如何确定组织能否使用集中部署，请参阅"确定外接程序的集中部署是否适用于[Microsoft 365 组织"。](/office365/admin/manage/centralized-deployment-of-add-ins)

> [!NOTE]
> 在未连接到 Microsoft 365 或部署面向 Office 2013 的 SharePoint 加载项或 Office 加载项本地环境中，请使用 [SharePoint 应用程序目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。 若要部署 COM/VSTO 加载项，请使用 ClickOnce 或 Windows Installer，如[部署 Office 解决方案](/visualstudio/vsto/deploying-an-office-solution)中所述。

## <a name="recommended-approach-for-deploying-office-add-ins"></a>部署 Office 加载项的推荐方法

请考虑分阶段部署 Office 加载项，以确保部署顺利进行。建议使用以下计划：

1. 为一小部分的企业利益干系人和 IT 部门成员部署加载项。 如果部署成功，则转到步骤 2。

2. 为企业内更多的将使用加载项的个人部署加载项。 如果部署成功，则转到步骤 3。

3. 为所有将使用加载项的个人部署加载项。

根据目标受众的规模，可能需要在此过程中添加步骤或删除步骤。

## <a name="publish-an-office-add-in-via-centralized-deployment"></a>通过集中部署发布 Office 加载项

开始之前，请确认组织满足使用集中部署的所有要求，如"确定外接程序的集中部署是否适用于 [Microsoft 365 组织"中所述](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)。

如果组织满足所有要求，请完成以下步骤以通过集中部署发布 Office 加载项：

1. 使用你的工作或教育帐户登录 Microsoft 365。
2. 选择左上角的应用启动器图标，然后选择“**管理员**”。
3. 在导航菜单中，按“**显示更多内容**”，然后选择“**设置**” > “**服务和加载项**”。
4. 如果在页面顶部看到一条消息，宣布推出新的 Microsoft 365 管理中心，请选择该消息以转到管理中心预览版 (请参阅"关于 [Microsoft 365](/microsoft-365/admin/admin-overview/about-the-admin-center) 管理中心") 。
5. 在页面顶部选择“**部署加载项**”。
6. 查看要求后，请选择“**下一步**”。
7. 在“**集中部署**”页面上，选择以下选项之一：

    - **我想从 Office 应用商店添加加载项。**
    - **我在此设备上具有清单文件 (.xml)。** 对于此选项，请选择“浏览”以找到想要使用的清单文件 (.xml)。
    - **我具有清单文件的 URL。** 对于此选项，请在提供的字段中键入清单的 URL。

    ![Microsoft 365 Add-In中的"新建"对话框](../images/new-add-in.png)

8. 如果选择了此选项以从 Office 应用商店添加某个加载项，请选择该加载项。 可以通过“**为你推荐**”、“**评级**”或“**名称**”类别，查看可用的加载项。 仅能从 Office 应用商店添加免费加载项。 目前不支持添加付费加载项。

    > [!NOTE]
    > 使用 Office 应用商店选项，无需干预，用户即可自动获得加载项的更新和增强功能。

    ![在 Microsoft 365 管理中心中选择加载项对话框](../images/select-an-add-in.png)

9. 查看 **外接程序** 详细信息、隐私策略和许可条款后，选择"继续"。

    ![Microsoft 365 管理中心中选定的外接程序页面](../images/selected-add-in-admin-center.png)

10. 在"**分配用户"** 页上，选择"**每个人****"、"特定用户/组"** 或"**仅我"。** 使用“搜索”框查找要向其部署加载项的用户和组。 对于 Outlook 外接程序，还可以选择部署方法 **"** 固定"、"**可用**"或"**可选"。**

    ![在 Microsoft 365 管理中心中管理有权访问和部署方法的用户](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > 利用 [SSO ](../develop/sso-in-office-add-ins.md) (单一登录) 将提示管理员同意加载项清单中列出的范围。  如果跨多个加载项使用同一支持服务 (则同一 Azure 应用 ID 用于不同外接程序) 中的 SSO，将提示每个外接程序的范围征得每个部署的同意。 此页面还将显示加载项所需的权限列表。

11. 完成后，选择"部署 **"。** 此过程可能最多用时 3 分钟。 然后，按“**下一步**”完成演练。 现在，可以看到此加载项与其他应用一起显示在 Office 365 中。

    > [!NOTE]
    > 当管理员 **选择"部署**"时，将给予所有用户同意。

    ![Microsoft 365 管理中心中的应用列表](../images/citations.png)

> [!TIP]
> 为组织中的用户和/或组部署新加载项时，请考虑向他们发送一封电子邮件，说明加载项的应用场景和使用方式，并添加相关帮助内容、FAQ 或其他支持资源的链接。

## <a name="considerations-when-granting-access-to-an-add-in"></a>授予加载项的访问权限时的注意事项

管理员可以将加载项分配给组织中的每个人或组织内的特定用户和/或组。 以下列表描述了每个选项的含义：

- **每个人**：顾名思义，此选项为租户中的每位用户分配加载项。请谨慎使用此选项，且仅应用于真正在组织中通用的加载项。

- **用户**：如果将加载项分配给单个用户，则每次要将其分配给其他用户时，都需要更新此加载项的集中部署设置。 同样，每次要删除某个用户对该加载项的访问权限时，都需要更新该加载项的集中部署设置。

- **组**：如果将加载项分配给组，则会自动为被添加到此组的用户分配此加载项。 同样，当将某个用户从组中删除时，此用户将自动失去对此加载项的访问权限。 在任一情况下，无需 Microsoft 365 管理员执行任何其他操作。

一般情况下，为了便于维护，我们建议尽可能使用组来分配加载项。 但是，在想要将加载项的访问权限限制在极少数用户的情况下，将加载项分配给特定用户的做法可能更为实用。

## <a name="add-in-states"></a>加载项状态

下表介绍了加载项的不同状态。

|状态|此状态如何出现|影响|
|-----|--------------------|------|
|**活动**|管理员已上传加载项并已将其分配给用户和/或组。|已为其分配加载项的用户和/或组，可在相关的 Office 客户端中找到它。|
|**已禁用**|管理员已禁用加载项。|已为其分配加载项的用户和/或组不再能够访问它。 如果加载项状态从“已禁用”**** 更改为“活动”****，则用户和组将重新获得对它的访问权限。|
|**已删除**|管理员已删除加载项。|已为其分配加载项的用户和/或组不再能够访问它。|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a>更新通过集中部署发布的 Office 加载项

通过集中部署发布 Office 外接程序后，在 Web 应用程序中实现这些更改后，对加载项 Web 应用程序进行的任何更改将自动提供给所有用户。 对外接程序 [的 XML](../develop/add-in-manifests.md) 清单文件所做的更改（例如，更新加载项的图标、文本或外接程序命令）发生如下：

- 业务线外接程序：如果管理员在通过 Microsoft 365 管理中心实现集中部署时，通过设备或通过指向 URL) 显式上载清单文件 (，则管理员必须上载包含所需更改的新清单文件。 上传更新后的清单文件后，加载项就会在下次相关 Office 应用启动时更新。

  > [!NOTE]
  > 管理员无需删除 LOB 加载项来进行更新。 在"加载项"部分，管理员只需选择 LOB 加载项，然后按右下角的"更新外接程序"按钮即可调用此功能。
  >
  > ![Screenshot shows the Update add-in dialog in Microsoft 365 admin center](../images/update-add-in-admin-center.png)

- **Office 应用商店** 外接程序：如果管理员在通过 Microsoft 365 管理中心实施集中部署时从 Office 应用商店选择了外接程序，并且 Office 应用商店中的外接程序更新，加载项稍后会通过集中部署进行更新。 可能需要 24 小时，应用商店外接程序更新才能针对所有最终用户流动。 在此持续时间后，下一次为这些用户重新启动相关 Office 应用程序时，外接程序将更新。 用户还可以通过选择"插入选项卡加载项管理托管选项卡点击刷新"来触发手动刷新，获取最新的应用商店  >    >    >  **外接程序版本**。

## <a name="end-user-experience-with-add-ins"></a>加载项最终用户体验

通过集中部署发布加载项后，最终用户可以在加载项支持的任何平台上开始使用它。

如果外接程序支持外接程序命令，则这些命令将出现在为其部署外接程序的所有用户的 Office 应用程序功能区上。 在以下的示例中，**搜索引文** 命令将显示在 **引文** 加载项的功能区上。

![Screenshot shows a section of the Office app ribbon with the Search Citation command highlighted in the Citations add-in](../images/search-citation.png)

如果加载项不支持加载项命令，用户可以通过执行以下操作将其添加到 Office 应用程序中：

1. 在 Word 2016 或更高版本、Excel 2016 或更高版本，或 PowerPoint 2016 或更高版本，选择“**插入**” > “**我的加载项**”。
2. 在加载项窗口中选择“**管理托管**”选项卡。
3. 选择加载项，然后选择“**添加**”。

    ![屏幕截图显示 Office 应用程序的“Office 加载项”页的“管理托管”选项卡。 引文加载项显示在此选项卡上。](../images/office-add-ins-admin-managed.png)

但是，对于 Outlook 2016 或更高版本，用户可以执行以下操作：

1. 在 Outlook 中，选择“**开始**” > “**应用商店**”。
2. 选择“加载项”选项卡下的“**管理员管理**”选项卡。
3. 选择加载项，然后选择“**添加**”。

    ![屏幕截图显示了 Outlook 应用程序的“应用商店”页面的管理员管理区域。](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a>另请参阅

- [确定加载项集中部署是否适用于 Microsoft 365 组织](/office365/admin/manage/centralized-deployment-of-add-ins)

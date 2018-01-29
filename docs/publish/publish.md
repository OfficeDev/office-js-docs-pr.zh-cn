# <a name="deploy-and-publish-your-office-add-in"></a>部署和发布 Office 外接程序

可以使用几种方法之一来部署 Office 外接程序，以用于对用户进行测试或分发：

|**方法**|**Use...**|
|:---------|:------------|
|[旁加载](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|在开发过程中测试在 Windows、Office Online、iPad 或 Mac 上运行的加载项。|
|[集中部署](centralized-deployment.md)|在云或混合部署中，使用 Office 365 管理中心将加载项分发给组织中的用户。|
|[SharePoint 目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|在本地环境中向组织中的用户分发加载项。|
|[Office 应用商店](https://dev.office.com/officestore/docs/submit-to-the-office-store)|用于向用户公开分发加载项。|
|[Exchange 服务器](#outlook-add-in-deployment)|在本地或联机环境中向用户分发 Outlook 加载项序。|

>**注意：**如果计划将加载项提交到 Office 应用商店，请务必遵循 [Office 应用商店验证策略](https://msdn.microsoft.com/zh-CN/library/jj220035.aspx)。 例如，加载项必须适用于支持你定义的方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://dev.office.com/officestore/docs/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性页](https://dev.office.com/add-in-availability)）。

## <a name="deployment-options-by-office-host"></a>Office 主机的部署选项

可用的部署选项具体取决于你定位的 Office 主机以及所创建的加载项的类型。

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Word、Excel 和 PowerPoint 加载项的部署选项

| 扩展点 | 旁加载 | Office 365 管理中心 |Office 应用商店| SharePoint 目录*  |
|:----------------|:------------|:-------------------|:--------------------------------|:-------------|
| 内容         | X           | X                  | X                               | X|
| 任务窗格       | X           | X                  | X                               | X|
| 命令         | X           | X                  | X                               |  |

&#42; SharePoint 目录不支持 Office 2016 for Mac。

### <a name="deployment-options-for-outlook-add-ins"></a>Outlook 外接程序的部署选项

| 扩展点 | 旁加载 | Exchange 服务器 | Office 应用商店 |
|:---------|:------------|:----------------|:-------------|
| 邮件应用 | X           | X               | X            |
| 命令  | X           | X               | X            |

## <a name="deployment-methods"></a>部署方法

以下各节提供了有关向组织中的用户分发 Office 加载项的最常用的部署方法的其他信息。

有关最终用户如何获取、插入和运行加载项的信息，请参阅[开始使用 Office 加载项](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)。

### <a name="centralized-deployment-via-the-office-365-admin-center"></a>通过 Office 365 管理中心进行集中部署 

通过 Office 365 管理中心，管理员可以轻松地为组织内的用户和组部署 Office 加载项。 通过管理中心部署加载项后，用户可立即在其 Office 应用程序中使用此加载项，而无需进行客户端配置。 可以通过集中部署来部署内部加载项以及 ISV 提供的加载项。

有关详细信息，请参阅[通过 Office 365 管理中心使用集中部署发布 Office 加载项](centralized-deployment.md)。

### <a name="sharepoint-catalog-deployment"></a>SharePoint 目录部署

SharePoint 加载项目录是特殊的网站集，可创建它来托管 Word、Excel 和 PowerPoint 加载项。由于 SharePoint 目录不支持在清单的 `VersionOverrides` 节点中实现的新加载项功能（包括加载项命令），因此建议尽可能通过管理中心进行集中部署。 通过 SharePoint 目录部署的加载项命令默认在任务窗格中打开。

如果要在本地环境中部署外接程序，请使用 SharePoint 目录。 有关详细信息，请参阅[将任务窗格和内容外接程序发布到 SharePoint 目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。

>**注意：**SharePoint 目录不支持 Office 2016 for Mac。 若要向 Mac 客户端部署 Office 加载项，必须将其提交到 [Office 应用商店]。 

### <a name="outlook-add-in-deployment"></a>Outlook 加载项部署

对于不使用 Azure AD 标识服务的本地和联机环境，可以通过 Exchange 服务器部署 Outlook 外接程序。 

Outlook 外接程序部署需要以下内容：

- Office 365、Exchange Online 或 Exchange Server 2013 或更高版本
- Outlook 2013 或更高版本

若要将外接程序分配给租户，请使用 Exchange 管理中心从文件或 URL直接上载清单，或从 Office 应用商店添加外接程序。 若要将外接程序分配给单个用户，则必须使用 Exchange PowerShell。 有关详细信息，请参阅 TechNet 上的[安装或删除组织的 Outlook 外接程序](https://technet.microsoft.com/zh-cn/library/jj943752(v=exchg.150).aspx)。

## <a name="additional-resources"></a>其他资源

- [部署和安装 Outlook 外接程序以进行测试](../outlook/testing-and-tips.md) 
- [提交至 Office 应用商店][Office 应用商店]
- [Office 外接程序的设计准则](../design/add-in-design)
- [创建有效的 Office 应用商店外接程序](https://msdn.microsoft.com/zh-CN/library/jj635874.aspx)
- [解决 Office 外接程序中的用户错误](../testing/testing-and-troubleshooting.md)

[Office 应用商店]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability

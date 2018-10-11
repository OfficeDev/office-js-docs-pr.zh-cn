# <a name="identity-api-requirement-sets"></a>Identity API 要求集

要求集就是已命名的 API 成员组。 Office 加载项使用清单中指定的要求集或使用运行检查，以确定 Office 主机是否支持加载项所需的 API。 有关更多信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Office 加载项运行跨多个Office版本。 下表列出标识 API 要求集，支持要求集的 Office 主机应用程序和 Office 应用程序的内部版本号或版本号。

|  要求集  | Office 2013 for Windows | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com 和 Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | 不适用 | 预览 ***** | 即将推出 | 预览 *****| 预览 | 预览| 即将推出 | 即将推出 |

> ***** 在预览阶段，在Windows 2016 和 Mac上仅对使用快速选项的内测计划用户提供标识 API 支持。 若要加入内测计划，请参阅 [加入 Office 内测](https://products.office.com/office-insider?tab=tab-1)。 要切换到快速跟踪，请参阅 [快速内测](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961)。

要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的 Office 版本？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

有关通用 API 要求集的信息，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="identityapi-11"></a>IdentityAPI 1.1 

单点登录 IdentityAPI 1.1 是 API 的第一个版本。 有关此 API 的详细信息，请参阅 [外接程序中启用 SSO](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) 的 [SSO API 参考](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)部分。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 加载项 XML 清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)

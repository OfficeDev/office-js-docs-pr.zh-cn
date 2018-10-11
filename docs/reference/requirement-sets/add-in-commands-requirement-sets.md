# <a name="add-in-commands-requirement-sets"></a>加载项命令要求集

要求集就是已命名的 API 成员组。 Office 加载项使用清单中指定的要求集或使用运行时检查，以确定 Office 主机是否支持加载项所需的 API。 有关更多信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

加载项命令是可扩展 Office UI 并在加载项中启动操作的 UI 元素。可以使用加载项命令在功能区上添加按钮或将某项添加到上下文菜单。有关更多信息，请参阅 [Excel、Word 和 PowerPoint 的加载项命令](https://docs.microsoft.com/office/dev/add-ins/design/add-in-commands)和 [Outlook 的加载项命令](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)。

加载项命令的初始版本没有相应的要求集（即，没有 AddinCommands 1.0 要求集）。下表列出了支持初始版本的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。  

| 发布   |  Office 2013 for Windows | Office 2016 for Windows（非订阅） | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| 加载项命令（初始版本，无要求集） | 不适用 | 16.0.4678.1000 *仅在 Outlook 中支持* |版本 1603（内部版本 6769.0000）或更新版本 | 不适用 | 15.33 或更新版本| 2016 年 1 月 | |

加载项命令 1.1 要求集介绍了[随文档自动打开任务窗格](https://docs.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document)的功能。

下表列出了加载项命令 1.1 要求集、支持此要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。 

|  要求集  |  Office 2013 for Windows | Office 2016 for Windows（非订阅） | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for iPad  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.1  | 不适用 | 16.0.4678.1000 *仅在 Outlook 中支持*  | 版本 1705（内部版本 8121.1000）或更新版本 | 不适用 | 15.34 或更新版本| 2017 年 5 月 | |

要详细了解版本、内部版本号和 Office 在线服务器，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office 在线服务器 概述](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

有关通用 API 要求集的信息，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 加载项 XML 清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)

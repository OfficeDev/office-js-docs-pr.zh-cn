---
title: 创建适用于撰写窗体的 Outlook 加载项
description: 了解有关适用于撰写窗体的 Outlook 加载项的方案和功能。
ms.date: 10/03/2022
ms.localizationpriority: high
ms.openlocfilehash: ef81b21eaa0bc63a5bf38757cb188e8850ade443
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467249"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a>创建适用于撰写窗体的 Outlook 加载项

可以创建撰写加载项，这些外接程序是在撰写窗体中激活的 Outlook 加载项。 与读取外接程序 (在用户查看消息或约会) 时在读取模式下激活的 Outlook 加载项相比，撰写加载项在以下用户方案中可用。

- 在撰写窗体中撰写新的邮件、会议请求或约会。

- 查看或编辑现有约会或用户是组织者的会议项目。

   > [!NOTE]
   > If the user is on the RTM release of Outlook 2013 and Exchange 2013 and is viewing a meeting item organized by the user, the user can find read add-ins available. Starting in the Office 2013 SP1 release, there's a change such that in the same scenario, only compose add-ins can activate and be available.

- 在单独的撰写窗体中撰写内嵌响应邮件或答复邮件。

- 编辑会议请求或会议项目答复（“接受”、“暂定”或“拒绝”）。

- 建议新的会议项目时间。

- 转发或答复会议请求或会议项目。

In each of these compose scenarios, any add-in command buttons defined by the add-in are shown. For older add-ins that do not implement add-in commands, users can choose **Office Add-ins** in the ribbon to open the add-in selection pane, and then choose and start a compose add-in. The following figure shows add-in commands in a compose form.

![显示 Outlook 撰写窗体，其中包含外接程序命令。](../images/compose-form-commands.png)

下图显示了外接程序选择窗格，该窗格由两个不实施外接程序命令的撰写外接程序组成，当用户在 Outlook 中撰写内嵌答复时将激活这两个撰写外接程序。

![为撰写项目激活的模板邮件应用。](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a>撰写模式下可用的外接程序的类型

撰写加载项作为[用于 Outlook 的加载项命令](add-in-commands-for-outlook.md)实现。 若要激活用于撰写电子邮件或会议答复的加载项，则加载项在清单中包括 [MessageComposeCommandSurface 扩展点元素](/javascript/api/manifest/extensionpoint#messagecomposecommandsurface)。 若要激活用于撰写或编辑用户是组织者的约会或会议的加载项，则加载项包括 [AppointmentOrganizerCommandSurface 扩展点元素](/javascript/api/manifest/extensionpoint#appointmentorganizercommandsurface)。

> [!NOTE]
> 为不支持加载项命令在包含在 [OfficeApp](/javascript/api/manifest/officeapp) 元素中的 [Rule](/javascript/api/manifest/rule) 元素使用[激活规则](activation-rules.md)的服务器或客户端开发的加载项。 除非加载项是为较早的客户端和服务器专门开发的，否则新加载项应使用加载项命令。

## <a name="api-features-available-to-compose-add-ins"></a>撰写加载项可用的 API 功能

- [在 Outlook 的撰写窗体中添加和删除项目附件](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [在 Outlook 的撰写窗体中获取和设置项目数据](get-and-set-item-data-in-a-compose-form.md)
- [在 Outlook 中撰写约会或邮件时获取、设置或添加收件人](get-set-or-add-recipients.md)
- [在 Outlook 中撰写约会或邮件时获取或设置主题](get-or-set-the-subject.md)
- [在 Outlook 中撰写约会或邮件时将数据插入到正文中](insert-data-in-the-body.md)
- [在 Outlook 中撰写约会时获取或设置位置](get-or-set-the-location-of-an-appointment.md)
- [在 Outlook 中撰写约会时获取或设置时间](get-or-set-the-time-of-an-appointment.md)

## <a name="see-also"></a>另请参阅

- [适用于 Office 的 Outlook 加载项入门](../quickstarts/outlook-quickstart.md)

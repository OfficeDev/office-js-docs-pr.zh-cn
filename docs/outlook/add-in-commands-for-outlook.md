---
title: Outlook 加载项命令
description: Outlook 加载项命令提供了通过添加按钮或下拉菜单从功能区启动特定加载项操作的方法。
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 7705c168077d2a704ff16b05bfb82416cd7f4154
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094027"
---
# <a name="add-in-commands-for-outlook"></a>适用于 Outlook 的外接程序命令

Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.

> [!NOTE]
> 加载项命令仅适用于 Windows 版 Outlook 2013 或更高版本、Mac 版 Outlook 2016 或更高版本、iOS 版 Outlook、Android 版 Outlook、适用于 Exchange 2016 或更高版本的 Outlook 网页版、适用于 Microsoft 365 的 Outlook 网页版和 Outlook.com。
>
> 需要安装以下三个更新，Outlook 2013 才支持加载项命令：
> - [2016 年 3 月 8 日发布的 Outlook 安全更新程序](https://support.microsoft.com/kb/3114829)
> - [2016 年 3 月 8 日发布的 Office 安全更新程序 (KB3114816)](https://support.microsoft.com/help/3114816/march-8,-2016,-update-for-office-2013-kb3114816)
> - [2016 年 3 月 8 日发布的 Office 安全更新程序 (KB3114828)](https://support.microsoft.com/help/3114828/march-8,-2016,-update-for-office-2013-kb3114828)
>
> 需要安装[累积更新 5](https://support.microsoft.com/help/4012106/cumulative-update-5-for-exchange-server-2016)，Exchange 2016 才支持加载项命令。

Add-in commands are only available for add-ins that do not use [ItemHasAttachment, ItemHasKnownEntity, or ItemHasRegularExpressionMatch rules](activation-rules.md) to limit the types of items they activate on. However, [contextual add-ins](contextual-outlook-add-ins.md) can present different commands depending on whether the currently selected item is a message or appointment, and can choose to appear in read or compose scenarios. Using add-in commands if possible is a [best practice](../concepts/add-in-development-best-practices.md).

## <a name="creating-the-add-in-command"></a>创建外接程序命令

Add-in commands are declared in the add-in manifest in the [VersionOverrides element](../reference/manifest/versionoverrides.md). This element is an addition to the manifest schema v1.1 that ensures backward compatibility. In a client that doesn't support `VersionOverrides`, existing add-ins will continue to function as they did without add-in commands.

`VersionOverrides` 清单条目会为加载项指定很多内容，如主机、要添加到功能区的控件的类型、文本、图标以及任何相关联的功能。

When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.notificationmessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.

Developers should define icons for all required sizes so that the add-in commands will adjust smoothly along with the ribbon. The required icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels for desktop, and 48 x 48 pixels, 32 x 32 pixels, and 25 x 25 pixels for mobile.

## <a name="how-do-add-in-commands-appear"></a>加载项命令的显示方式

An add-in command appears on the ribbon as a button. When a user installs an add-in, its commands appear in the UI as a group of buttons. This can either be on the ribbon's default tab or on a custom tab. For messages, the default is either the **Home** or **Message** tab. For the calendar, the default is the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tab. For module extensions, the default is a custom tab. On the default tab, each add-in can have one ribbon group with up to 6 commands. On custom tabs, the add-in can have up to 10 groups, each with 6 commands. Add-ins are limited to only one custom tab.

当功能区变得拥挤时，加载项命令将显示在溢出菜单中。 用于加载项的加载项命令通常将组合在一起。

![功能区上的加载项命令按钮](../images/commands-normal.png)

![功能区和溢出菜单中的加载项命令按钮](../images/commands-collapsed.png)

When an add-in command is added to an add-in, the add-in name is removed from the app bar. Only the add-in command button on the ribbon remains.

### <a name="modern-outlook-on-the-web"></a>新式 Outlook 网页版

在 Outlook 网页版中，加载项名称显示在溢出菜单中。 如果加载项具有多个加载项命令，则可展开加载项菜单以查看一组标记有加载项名称的按钮。

![可在其中找到加载项命令按钮的溢出菜单](../images/commands-overflow-menu-web.png)

![显示加载项命令按钮的溢出菜单](../images/commands-overflow-menu-expand-web.png)

## <a name="what-ux-shapes-exist-for-add-in-commands"></a>外接程序命令存在哪些 UX 形状？

The UX shape for an add-in command consists of a ribbon tab in the host application that contains buttons that can perform various functions. Currently, three UI shapes are supported:

- 一个可执行 JavaScript 函数的按钮
- 一个启动任务窗格的按钮
- 显示另外两种类型的一个或多个按钮的下拉菜单的按钮

### <a name="executing-a-javascript-function"></a>执行 JavaScript 函数

Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service.

在模块扩展中，外接程序命令按钮可以执行与主要用户界面的内容交互的 JavaScript 函数。

![用于执行 Outlook 功能区上的功能的按钮。](../images/commands-uiless-button-1.png)

### <a name="launching-a-task-pane"></a>启动任务窗格

Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields.

The default width of the vertical task pane is 320 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.

![用于在 Outlook 功能区上打开任务窗格的按钮。](../images/commands-task-pane-button-1.png)

<br/>

This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. By default, this pane will not persist across messages. Add-ins can [support pinning](pinnable-taskpane.md) for the task pane and receive events when a new message is selected. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.

If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.

### <a name="drop-down-menu"></a>下拉菜单

A drop-down menu add-in command defines a static list of buttons. The buttons within the menu can be any mix of buttons that execute a function or buttons that open a task pane. Submenus are not supported.

![用于下拉 Outlook 功能区上的菜单的按钮。](../images/commands-menu-button-1.png)

## <a name="where-do-add-in-commands-appear-in-the-ui"></a>外接程序命令显示在 UI 中的什么位置？

以下四种方案支持外接程序命令：

### <a name="reading-a-message"></a>阅读邮件

用户在阅读邮件时，如果在阅读窗格或“**邮件**”选项卡的弹出式阅读表单中查看邮件，添加到默认选项卡的外接程序命令将出现在“**主页**”选项卡上。

### <a name="composing-a-message"></a>撰写邮件

用户在撰写邮件时，添加到默认选项卡的加载项命令将出现在“邮件”**** 选项卡上。

### <a name="creating-or-viewing-an-appointment-or-meeting-as-the-organizer"></a>以组织者的身份创建或查看约会或会议

When creating or viewing an appointment or meeting as the organizer, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tabs on pop-out forms. However, if the user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon.

### <a name="viewing-a-meeting-as-an-attendee"></a>以参与者的身份查看会议

When viewing a meeting as an attendee, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, or **Meeting Series** tabs on pop-out forms. However, if a user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon

### <a name="using-a-module-extension"></a>使用模块扩展

使用模块扩展时，加载项命令显示在扩展的自定义选项卡上。

## <a name="see-also"></a>另请参阅

- [加载项命令演示 Outlook 加载项](https://github.com/officedev/outlook-add-in-command-demo)
- [在清单中创建 Excel、PowerPoint 和 Word 加载项命令](../develop/create-addin-commands.md)

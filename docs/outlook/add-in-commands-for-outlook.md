---
title: Outlook 加载项命令
description: Outlook 加载项命令提供了通过添加按钮或下拉菜单从功能区启动特定加载项操作的方法。
ms.date: 10/11/2022
ms.localizationpriority: high
ms.openlocfilehash: d029fd4acc1a32c912c73d6e5f468b9c217b9262
ms.sourcegitcommit: 787fbe4d4a5462ff6679ad7fd00748bf07391610
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2022
ms.locfileid: "68546457"
---
# <a name="add-in-commands-for-outlook"></a>适用于 Outlook 的外接程序命令

Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.

> [!NOTE]
> 加载项命令仅适用于 Windows 版 Outlook 2013 或更高版本、Mac 版 Outlook 2016 或更高版本、iOS 版 Outlook、Android 版 Outlook、适用于 Exchange 2016 或更高版本的 Outlook 网页版、适用于 Microsoft 365 的 Outlook 网页版和 Outlook.com。
>
> 需要安装以下三个更新，Outlook 2013 才支持加载项命令：
> - [2016 年 3 月 8 日发布的 Outlook 安全更新程序](https://support.microsoft.com/kb/3114829)
> - [2016 年 3 月 8 日发布的 Office 安全更新程序 (KB3114816)](https://support.microsoft.com/topic/3d3eb171-78c2-0e61-62a2-85723bc4bcc0)
> - [2016 年 3 月 8 日发布的 Office 安全更新程序 (KB3114828)](https://support.microsoft.com/topic/54437016-d1e0-7aac-dbb7-4ecfbd57f5f0)
>
> 需要安装[累积更新 5](https://support.microsoft.com/topic/d67d7693-96a4-fb6e-b60b-e64984e267bd)，Exchange 2016 才支持加载项命令。

> [!TIP]
> 如果外接程序使用 XML 清单，则外接程序命令仅适用于不使用 [ItemHasAttachment、ItemHasKnownEntity 或 ItemHasRegularExpressionMatch 规则](activation-rules.md) 的加载项，以限制其激活的项类型。 但是， [上下文加载项](contextual-outlook-add-ins.md) 可以显示不同的命令，具体取决于当前选定的项是消息还是约会，并且可以选择出现在读取或撰写方案中。 如可能，使用外接程序命令将是[最佳做法](../concepts/add-in-development-best-practices.md)。

## <a name="create-the-ui-for-the-add-in-command"></a>为加载项命令创建 UI

外接程序命令在加载项清单中声明。 标记取决于清单的类型。

# <a name="xml-manifest"></a>[XML 清单](#tab/xmlmanifest)

外接程序命令在 [VersionOverrides 元素](/javascript/api/manifest/versionoverrides)中声明。 此元素是 XML 清单架构 v1.1 的一个补充，可确保向后兼容性。 在不支持 **\<VersionOverrides\>** 的环境中，现有的加载项将照常像在没有加载项命令的情况下正常运行。

**\<VersionOverrides\>** 清单条目为加载项指定许多内容，如应用程序、要添加到功能区的控件的类型、文本、图标以及任何关联的功能。

When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.notificationmessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.

# <a name="teams-manifest-developer-preview"></a>[Teams 清单 (开发人员预览) ](#tab/jsonmanifest)

外接程序命令使用“extensions.runtimes”和“extensions.ribbons”属性进行声明。 这些属性为加载项指定许多内容，例如应用程序、要添加到功能区中的控件类型、文本、图标和任何关联函数。

当外接程序需要提供状态更新（例如进度指示器或错误消息）时，它必须通过 [通知 API](/javascript/api/outlook/office.notificationmessages) 来执行此操作。 通知的处理还必须在清单的“runtimes.code.page”属性中指定的单独 HTML 文件中定义。

---
### <a name="icons"></a>图标

Developers should define icons for all required sizes so that the add-in commands will adjust smoothly along with the ribbon. The required icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels for desktop, and 48 x 48 pixels, 32 x 32 pixels, and 25 x 25 pixels for mobile.

## <a name="how-do-add-in-commands-appear"></a>加载项命令的显示方式

加载项命令在功能区中显示为按钮，或在下拉菜单中显示为菜单项。 当用户安装加载项时，其命令将作为一组按钮显示在 UI 中。 这可能出现在功能区的默认选项卡上，也可能出现在自定义选项卡上。对于消息，默认为“**开始**”或“**消息**”选项卡。对于日历，则默认为“**会议**”、“**会议事件**”、“**会议系列**”或“**约会**”选项卡。对于模块扩展，默认为自定义选项卡。在默认选项卡上，每个加载项可以具有一个功能区组，最多包含 6 个命令。 在自定义选项卡上，外接程序最多具有 10 个组，每个组具有 6 个命令。 外接程序限定为仅一个自定义选项卡。

当功能区变得拥挤时，加载项命令将显示在溢出菜单中。 用于加载项的加载项命令通常将组合在一起。

![功能区上的加载项命令按钮。](../images/commands-normal.png)

![功能区和溢出菜单中的加载项命令按钮。](../images/commands-collapsed.png)

When an add-in command is added to an add-in, the add-in name is removed from the app bar. Only the add-in command button on the ribbon remains.

### <a name="modern-outlook-on-the-web"></a>新式 Outlook 网页版

在 Outlook 网页版中，加载项名称显示在溢出菜单中。 如果加载项具有多个加载项命令，则可展开加载项菜单以查看一组标记有加载项名称的按钮。

![可在其中找到加载项命令按钮的溢出菜单。](../images/commands-overflow-menu-web.png)

![显示加载项命令按钮的溢出菜单。](../images/commands-overflow-menu-expand-web.png)

## <a name="what-are-the-types-of-add-in-commands"></a>加载项命令的类型是什么？

加载项命令的 UI 包括功能区按钮或下拉菜单中的项。 根据命令触发的操作类型，有两种类型的加载项命令。

- **任务窗格命令**：按钮或菜单项将打开加载项的任务窗格。 在清单中添加带有标记的此类加载项命令。 “代码隐藏”命令由 Office 提供。
- **函数命令**：按钮或菜单项运行任意 JavaScript。 代码几乎总是在 Office JavaScript 库中调用 API，但并非必须如此。 此类型的加载项通常不显示按钮或菜单项本身以外的 UI。 请注意以下有关函数命令的内容：

   - 触发的函数可以调用 [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) 方法来显示对话框，这是显示错误、显示进度或提示用户输入的好方法。
   - 函数命令运行的运行时是基于 [浏览器的完整运行时](../testing/runtimes.md#browser-runtime)。 它可以呈现 HTML 并调用 Internet 以发送或获取数据。

### <a name="run-a-function-command"></a>运行函数命令

Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service.

在模块扩展中，外接程序命令按钮可以执行与主要用户界面的内容交互的 JavaScript 函数。

![用于执行 Outlook 功能区上的功能的按钮。](../images/commands-uiless-button-1.png)

### <a name="launch-a-task-pane"></a>启动任务窗格

Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields.

The default width of the vertical task pane is 320 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.

![用于在 Outlook 功能区上打开任务窗格的按钮。](../images/commands-task-pane-button-1.png)

<br/>

This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. By default, this pane will not persist across messages. Add-ins can [support pinning](pinnable-taskpane.md) for the task pane and receive events when a new message is selected. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.

If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.

### <a name="drop-down-menu"></a>下拉菜单

下拉菜单加载项命令定义静态的项目列表。 菜单可以是执行函数或打开任务窗格的任何项组合。 不支持子菜单。

![用于下拉 Outlook 功能区上的菜单的按钮。](../images/commands-menu-button-1.png)

## <a name="where-do-add-in-commands-appear-in-the-ui"></a>外接程序命令显示在 UI 中的什么位置？

以下四种方案支持外接程序命令：

### <a name="reading-a-message"></a>阅读邮件

用户在阅读邮件时，如果在阅读窗格或“**邮件**”选项卡的弹出式阅读表单中查看邮件，添加到默认选项卡的外接程序命令将出现在“**主页**”选项卡上。

### <a name="composing-a-message"></a>撰写邮件

用户在撰写邮件时，添加到默认选项卡的加载项命令将出现在“邮件”选项卡上。

### <a name="creating-or-viewing-an-appointment-or-meeting-as-the-organizer"></a>以组织者的身份创建或查看约会或会议

When creating or viewing an appointment or meeting as the organizer, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tabs on pop-out forms. However, if the user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon.

### <a name="viewing-a-meeting-as-an-attendee"></a>以参与者的身份查看会议

When viewing a meeting as an attendee, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, or **Meeting Series** tabs on pop-out forms. However, if a user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon

### <a name="using-a-module-extension"></a>使用模块扩展

使用模块扩展时，加载项命令显示在扩展的自定义选项卡上。

## <a name="see-also"></a>另请参阅

- [加载项命令演示 Outlook 加载项](https://github.com/officedev/outlook-add-in-command-demo)
- [在清单中创建 Excel、PowerPoint 和 Word 加载项命令](../develop/create-addin-commands.md)
- [Outlook 加载项中的调试函数命令](debug-ui-less.md)
- [教程：生成邮件撰写 Outlook 外接程序](../tutorials/outlook-tutorial.md)

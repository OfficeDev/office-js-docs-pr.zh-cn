---
title: 加载项命令的基本概念
description: 了解如何将自定义功能区按钮和菜单项添加到 Office 作为 Office 加载项的一部分。
ms.date: 05/12/2020
localization_priority: Priority
ms.openlocfilehash: 2fe14a41c93b53164ab0fa3a7d25f5b9810b9c6a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093873"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a>Excel、PowerPoint 和 Word 的加载项命令

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.

有关此功能的概述，请观看视频 [Office 应用功能区中的加载项命令](https://channel9.msdn.com/events/Build/2016/P551)。

> [!NOTE]
> SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.

> [!IMPORTANT]
> Outlook 中也支持加载项命令。 有关详细信息，请参阅[适用于 Outlook 的加载项命令](../outlook/add-in-commands-for-outlook.md)。

*图 1：在 Excel Desktop 中运行命令的加载项*

![Excel 中的加载项命令屏幕截图](../images/add-in-commands-1.png)

*图 2：在 Excel 网页版中运行命令的加载项*

![Excel 网页版中加载项命令的屏幕截图](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a>命令功能

目前支持下列命令功能。

> [!NOTE]
> 内容加载项暂不支持加载项命令。

### <a name="extension-points"></a>扩展点

- 功能区选项卡 - 扩展内置选项卡或新建自定义选项卡。
- 上下文菜单 - 扩展所选上下文菜单。

### <a name="control-types"></a>控件类型

- 简单按钮 - 触发特定操作。
- 菜单 - 简单的下拉菜单，内含可触发操作的按钮。

### <a name="actions"></a>操作

- ShowTaskpane - 显示一个或多个在其中加载自定义 HTML 页的窗格。
- ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.  

### <a name="default-enabled-or-disabled-status-preview"></a>默认启用或禁用状态（预览版）

可指定在加载项启动时是启用还是禁用该命令，并以编程方式更改设置。

> [!NOTE]
> 此功能处于预览状态，并非在所有主机或方案中均受支持。 有关详细信息，请参阅[启用和禁用加载项命令](disable-add-in-commands.md)。

## <a name="supported-platforms"></a>支持的平台

目前，以下平台支持加载项命令。

- Windows 版 Office（内部版本 16.0.6769 及更高版本，关联至 Microsoft 365 订阅）
- Windows 版 Office 2019
- Mac 版 Office（内部版本 15.33 及更高版本，关联至 Microsoft 365 订阅）
- Mac 版 Office 2019
- Office 网页版

> [!NOTE]
> 有关 Outlook 支持的信息，请参阅[适用于 Outlook 的加载项命令](../outlook/add-in-commands-for-outlook.md)。

## <a name="debugging"></a>调试

必须在 Office 网页版中运行加载项命令，才能调试命令。 有关详细信息，请参阅[在 Office 网页版中调试加载项](../testing/debug-add-ins-in-office-online.md)。

## <a name="best-practices"></a>最佳做法

在开发外接程序命令时应用下面的最佳做法：

- Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.
- Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.
- 关于命令在 Office 应用功能区中的位置：
    - Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).
    - Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).  
    - 如果你拥有 6 个以上的顶级命令命令，将命令放置在自定义选项卡上。
    - Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.
    - 请勿添加不必要的按钮，这样会增加加载项占用的空间。

     > [!NOTE]
     > 占用过多空间的加载项可能无法通过 [AppSource 验证](/legal/marketplace/certification-policies)。

- 对于所有图标，请遵循[图标设计准则](add-in-icons.md)。
- 提供也可以在不支持命令的主机上运行的加载项的版本。 单个加载项清单可以在命令感知（带有命令）和非命令感知（作为任务窗格）的主机中工作。

   *图 3. Office 2013 中的任务窗格加载项，以及 Office 2016 中使用加载项命令的相同加载项*

   ![显示 Office 2013 中的任务窗格加载项，以及 Office 2016 中使用加载项命令的相同加载项的屏幕截图](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a>后续步骤

加载项命令的最佳入门方式是参照 GitHub 上的 [Office 加载项命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)。

若要详细了解如何在清单中指定加载项命令，请参阅[在清单中创建加载项命令](../develop/create-addin-commands.md)和 [VersionOverrides](../reference/manifest/versionoverrides.md) 参考内容。

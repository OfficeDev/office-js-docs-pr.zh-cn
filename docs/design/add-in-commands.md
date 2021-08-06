---
title: 加载项命令的基本概念
description: 了解如何将自定义功能区按钮和菜单项添加到 Office 作为 Office 加载项的一部分。
ms.date: 07/27/2021
localization_priority: Priority
ms.openlocfilehash: 4ee2e53a1d2a74a2663a372aeb080c5f32da1bde
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773207"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a>Excel、PowerPoint 和 Word 的加载项命令

外接程序命令是 UI 元素，可扩展 Office UI，并在外接程序中启动操作。使用外接程序命令，可以在功能区上添加按钮，也可以向上下文菜单添加项。当用户选择外接程序命令时，将启动操作，如运行 JavaScript 代码或在任务窗格中显示外接程序页面。外接程序命令可帮助用户查找和使用外接程序，从而提高外接程序的采用率和重用率以及客户保留率。

有关此功能的概述，请观看视频 [Office 应用功能区中的加载项命令](https://channel9.msdn.com/events/Build/2016/P551)。

> [!NOTE]
> SharePoint 目录不支持加载项命令。 可以通过[集中部署](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)或 [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) 部署加载项命令，也可以使用[旁加载](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)部署加载项命令以供测试。

> [!IMPORTANT]
> Outlook 中也支持加载项命令。 有关详细信息，请参阅[适用于 Outlook 的加载项命令](../outlook/add-in-commands-for-outlook.md)。

*图 1：在 Excel Desktop 中运行命令的加载项*

![显示 Excel 功能区中突出显示的加载项命令屏幕截图。](../images/add-in-commands-1.png)

*图 2：在 Excel 网页版中运行命令的加载项*

![显示 Excel 网页版中加载项命令的屏幕截图。](../images/add-in-commands-2.png)

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
- ExecuteFunction - 加载一个不可见的 HTML 页，然后在其中执行一个 JavaScript 函数。若要在你的函数（例如错误、进度或其他输入）中显示 UI，你可以使用 [displayDialog](/javascript/api/office/office.ui) API。  

### <a name="default-enabled-or-disabled-status"></a>默认启用或禁用状态

可指定在加载项启动时是启用还是禁用该命令，并以编程方式更改设置。

> [!NOTE]
> 此功能并非在所有 Office 应用程序或方案中受到支持。 有关详细信息，请参阅[启用和禁用加载项命令](disable-add-in-commands.md)。

### <a name="position-on-the-ribbon-preview"></a>功能区上的位置（预览）

可以指定自定义选项卡在 Office 应用程序功能区上的显示位置，例如“在“主页”选项卡右侧”。

> [!NOTE]
> 并非所有 Office 应用程序或方案均支持此功能。 有关详细信息，请参阅[在功能区上定位自定义选项卡](custom-tab-placement.md)。

### <a name="integration-of-built-in-office-buttons-preview"></a>内置 Office 按钮集成（预览）

可将内置的 Office 功能区按钮插入到自定义命令组和自定义功能区选项卡中。

> [!NOTE]
> 并非所有 Office 应用程序或方案均支持此功能。 有关详细信息，请参阅[将内置 Office 按钮集成到自定义选项卡中](built-in-button-integration.md)。

### <a name="contextual-tabs-preview"></a>上下文选项卡（预览）

可指定一个选项卡在某些情况下只在功能区中可见，例如在Excel中选择图表时。

> [!NOTE]
> 并非所有 Office 应用程序或方案均支持此功能。 更多信息，请参见[在Office插件中创建自定义上下文选项卡](contextual-tabs.md)。

## <a name="supported-platforms"></a>支持的平台

目前，以下平台支持加载项命令，但先前[命令功能](#command-capabilities)的小节中指定的限制除外。

- Windows 版 Office（内部版本 16.0.6769 及更高版本，关联至 Microsoft 365 订阅）
- Windows 版 Office 2019
- Mac 版 Office（内部版本 15.33 及更高版本，关联至 Microsoft 365 订阅）
- Mac 版 Office 2019
- Office 网页版

> [!NOTE]
> 有关 Outlook 支持的信息，请参阅[适用于 Outlook 的加载项命令](../outlook/add-in-commands-for-outlook.md)。

## <a name="debug"></a>调试

必须在 Office 网页版中运行加载项命令，才能调试命令。 有关详细信息，请参阅[在 Office 网页版中调试加载项](../testing/debug-add-ins-in-office-online.md)。

## <a name="best-practices"></a>最佳做法

在开发外接程序命令时应用下面的最佳做法。

- 使用命令来表示会给用户带来明确具体结果的特定操作。不要在单个按钮中组合多个操作。
- 提供使您的外接程序中的常见任务执行效率更高的具体操作。尽量减少完成一个操作的步骤。
- 关于命令在 Office 应用功能区中的位置：
  - 将命令放置在现有的选项卡（插入、审阅等）上，如果提供的功能适合那个位置。例如，如果外接程序允许用户插入媒体，则将组添加到“插入”选项卡。请注意，并非所有选项卡都在所有的 Office 版本之间可用。有关详细信息，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。
  - 如果此功能不适合其他选项卡，且顶级命令少于 6 个，请将命令置于“开始”选项卡中。此外，如果加载项需要跨 Office 版本（如 Office 网页版或 Office 桌面版）运行，且并非所有版本都有相应选项卡（例如，Office 网页版中没有“设计”选项卡），也可以将命令添加到“开始”选项卡中。  
  - 如果你拥有 6 个以上的顶级命令命令，将命令放置在自定义选项卡上。
  - 对组进行命名以与外接程序的名称相匹配。如果你拥有多个组，则基于对应组中的命令提供的功能为每个组命名。
  - 请勿添加不必要的按钮，这样会增加加载项占用的空间。
  - 请勿要将“自定义”选项卡置于“主页”选项卡左侧，也不要在打开文档时默认将其放在焦点上，除非加载项是用户与文档进行交互的主要方式。 过分强调加载项的不便，并惹恼用户和管理员。
  - 如果加载项是用户与文档进行交互的主要方式，而且你具有自定义的功能区选项卡，请考虑将用户经常需要的 Office 功能按钮集成到该选项卡中。
  - 如果用自定义标签提供的功能只能在特定的上下文中使用，请使用[自定义下文选项卡](contextual-tabs.md)。 如果使用自定义上下文选项卡，请确保[在插件运行在不支持自定义上下文标签的平台上时，实行后退体验](contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

  > [!NOTE]
  > 占用过多空间的加载项可能无法通过 [AppSource 验证](/legal/marketplace/certification-policies)。

- 对于所有图标，请遵循[图标设计准则](add-in-icons.md)。
- 提供也可以在不支持命令的 Office 应用程序上运行的加载项版本。单个加载项清单可以在命令感知型（带有命令）和非命令感知（作为任务窗格）型应用程序中工作。

   *图 3. Office 2013 中的任务窗格加载项，以及 Office 2016 中使用加载项命令的相同加载项*

   ![比较 Office 2013 中的任务窗格加载项和 Office 2016 中使用加载项命令的相同加载项的屏幕截图。 在 2013 版本中，任务窗格必须包含所有命令，而在 2016 版本中，命令可以位于功能区中。](../images/office-task-pane-add-ins.png)

## <a name="next-steps"></a>后续步骤

加载项命令的最佳入门方式是参照 GitHub 上的 [Office 加载项命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)。

若要详细了解如何在清单中指定加载项命令，请参阅[在清单中创建加载项命令](../develop/create-addin-commands.md)和 [VersionOverrides](../reference/manifest/versionoverrides.md) 参考内容。

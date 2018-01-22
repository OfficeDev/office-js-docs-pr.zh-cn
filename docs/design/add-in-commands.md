
# <a name="add-in-commands-for-excel-word-and-powerpoint"></a>Excel、Word 和 PowerPoint 的外接程序命令

外接程序命令是 UI 元素，可扩展 Office UI，并在外接程序中启动操作。使用外接程序命令，可以在功能区上添加按钮，也可以向上下文菜单添加项。当用户选择外接程序命令时，将启动操作，如运行 JavaScript 代码或在任务窗格中显示外接程序页面。外接程序命令可帮助用户查找和使用外接程序，从而提高外接程序的采用率和重用率以及客户保留率。

有关此功能的概述，请观看视频 [Office 功能区中的加载项命令](https://channel9.msdn.com/events/Build/2016/P551)。

>**注意：**SharePoint 目录不支持加载项命令。 可以通过[集中部署](../publish/centralized-deployment.md)或 [Office 应用商店](https://dev.office.com/officestore/docs/submit-to-the-office-store)部署加载项命令，也可以使用[旁加载](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)部署加载项命令来进行测试。 

**命令在 Excel Desktop 中运行的加载项**

![Excel 中的外接程序命令屏幕截图](../images/addincommands1.png)

**命令在 Excel Online 中运行的外接程序**

![Excel Online 中的外接程序命令屏幕截图](../images/addincommands2.png)

## <a name="command-capabilities"></a>命令功能
目前支持下列命令功能。

> **注意：**内容外接程序当前不支持外接程序命令。

**扩展点**

- 功能区选项卡 - 扩展内置选项卡或新建自定义选项卡。
- 上下文菜单 - 扩展选定上下文菜单。 

**控件类型**

- 简单按钮 - 触发特定操作。
- 菜单 - 简单的下拉菜单，内含可触发操作的按钮。

**操作**

- ShowTaskpane - 显示一个或多个在其中加载自定义 HTML 页的窗格。
- ExecuteFunction - 加载一个不可见的 HTML 页，然后在其中执行一个 JavaScript 函数。若要在你的函数（例如错误、进度或其他输入）中显示 UI，你可以使用 [displayDialog](http://dev.office.com/reference/add-ins/shared/officeui) API。  

## <a name="supported-platforms"></a>支持的平台
目前，以下平台支持外接程序命令：

- Office for Windows Desktop 2016（内部版本 16.0.6769+）
- Office for Mac（内部版本 15.33+）
- Office Online 

即将推出更多受支持的平台。

## <a name="best-practices"></a>最佳做法

在开发外接程序命令时应用下面的最佳做法：

- 使用命令来表示会给用户带来明确具体结果的特定操作。不要在单个按钮中组合多个操作。
- 提供使您的外接程序中的常见任务执行效率更高的具体操作。尽量减少完成一个操作的步骤。
- 关于命令在 Office 功能区中的位置：
    - 将命令放置在现有的选项卡（插入、审阅等）上，如果提供的功能适合那个位置。例如，如果外接程序允许用户插入媒体，则将组添加到“插入”选项卡。请注意，并非所有选项卡都在所有的 Office 版本之间可用。有关详细信息，请参阅 [Office 外接程序 XML 清单](../overview/add-in-manifests.md)。 
    - 如果功能不适合其他选项卡，并且你拥有少于 6 个的顶级命令，将命令放置在“开始”选项卡上。如果外接程序需要跨 Office 版本（如 Office Desktop 和 Office Online）运行，并且某个选项卡并非在所有版本中（例如，“设计”选项卡不存在于 Office Online 中）都提供，你也可以将命令添加到“开始”选项卡。  
    - 如果你拥有 6 个以上的顶级命令命令，将命令放置在自定义选项卡上。 
    - 对组进行命名以与外接程序的名称相匹配。如果你拥有多个组，则基于对应组中的命令提供的功能为每个组命名。
    - 不要添加不必要的按钮，从而为你的外接程序留出更多的空间。

     >**注意：**占用过多空间的外接程序可能无法通过 [Office 应用商店验证](https://dev.office.com/officestore/docs/validation-policies)。

- 对于所有图标，请遵循[图标设计准则](../design/design-icons.md)。
- 提供也可以在不支持命令的主机上运行的外接程序的版本。单个外接程序清单可以在命令感知（带有命令）和非命令感知（作为任务窗格）的主机中工作。

    ![显示 Office 2013 中的任务窗格外接程序，以及 Office 2016 中使用外接程序命令的相同外接程序的屏幕截图](../images/4f90a3cc-8cc4-4879-9a03-0bb2b6079026.png)


## <a name="next-steps-to-get-started"></a>入门所需的后续步骤

外接程序命令的最佳入门方式是参照 GitHub 上的 [Office 外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)。

若要详细了解如何在清单中指定外接程序命令，请参阅[在清单中定义外接程序命令](../develop/define-add-in-commands.md)和 [VersionOverrides](http://dev.office.com/reference/add-ins/manifest/versionoverrides) 参考内容。






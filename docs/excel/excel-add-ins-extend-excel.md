# <a name="extend-excel-functionality"></a>扩展 Excel 功能

除了与工作簿中的内容进行交互，Excel 加载项还可以添加自定义功能区按钮或菜单命令、插入任务窗格、打开对话框，甚至可以将基于 Web 的丰富内容直接嵌入到工作表中。

## <a name="add-in-commands"></a>加载项命令

加载项命令是 UI 元素，可扩展 Excel UI，并在加载项中启动操作。 使用加载项命令，可以在功能区上添加按钮，也可以向 Excel 中的上下文菜单添加项。 当用户选择加载项命令时，将启动操作，如运行 JavaScript 代码或在任务窗格中显示加载项页面。 

**加载项命令**

![Excel 中的加载项命令](../images/Excel_add-in_commands_Script-Lab.png)

有关命令功能、受支持的平台和开发加载项命令第最佳做法的详细信息，请参阅[适用于 Excel、Word 和 Powerpoint 的加载项命令](../design/add-in-commands.md)。

## <a name="task-panes"></a>任务窗格

任务窗格作为界面图面，通常出现在 Excel 内的窗口右侧。 任务窗格允许用户访问界面控件，此类控件运行代码以修改 Excel 文档或显示数据源中的数据。 

**任务窗格**

![Excel 中的任务窗格加载项](../images/Excel_add-in_task_pane_Insights.png)

有关任务窗格的详细信息，请参阅 [Office 加载项中的任务窗格](../design/task-pane-add-ins.md)。有关在 Excel 中实现任务窗格的示例，请参阅 [Excel 加载项 JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)。

## <a name="dialog-boxes"></a>对话框

对话框是浮动在活动的 Excel 应用程序窗口之上的界面。 可以将对话框用于以下任务，如显示无法直接在任务窗格中打开的登录页、请求用户确认操作，或托管如果局限在任务窗格中可能过小的视频。 若要在 Excel 加载项中打开对话框，请使用[对话框 API](http://dev.office.com/reference/add-ins/shared/officeui)。

**对话框**

![Excel 中的加载项对话框](../images/Excel_add-in_dialog_choose-number.png)

有关对话框和对话框 API 的详细信息，请参阅 [Office 加载项中的对话框](../design/dialog-boxes.md)和[在 Office 加载项中使用对话框 API](../develop/dialog-api-in-office-add-ins.md)。

## <a name="content-add-ins"></a>内容加载项

内容加载项是可以直接嵌入到 Excel 文档中的图面。 可以使用内容加载项在工作表中嵌入基于 Web 的丰富对象，如图表、数据可视化效果或媒体，或为用户提供对界面控件的访问权限，这些控件运行代码以修改 Excel 文档，或显示来自数据源的数据。 在你要将功能直接嵌入文档时，请使用内容加载项。

**内容加载项**

![Excel 中的内容加载项](../images/Excel_add-in_content_map.png)

有关内容加载项的详细信息，请参阅 [Office 内容加载项](../design/content-add-ins.md)。有关在 Excel 中实现内容加载项的示例，请参阅 GitHub 中的 [ Excel 内容加载项 Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)。

## <a name="additional-resources"></a>其他资源

- [Excel、Word 和 PowerPoint 的加载项命令](../design/add-in-commands.md)
- [在清单中定义加载项命令](../develop/define-add-in-commands.md)
- [Github 上的 Office 加载项命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)
- [Office 加载项中的任务窗格](../design/task-pane-add-ins.md)
- [Excel 加载项：JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [Office 加载项中的对话框](../design/dialog-boxes.md)
- [在 Office 加载项中使用对话框 API](../develop/dialog-api-in-office-add-ins.md)
- [Office 加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [内容 Office 加载项](../design/content-add-ins.md)
- [Excel 内容加载项：Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)

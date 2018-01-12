
# <a name="use-office-ui-fabric-js-in-office-add-ins"></a>在 Office 外接程序中使用 Office UI Fabric JS

Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。如果仅使用 JavaScript，而不使用 Angular 或 React 等框架，可考虑使用 Fabric JS 创建用户体验。有关详细信息，请参阅 [Office UI Fabric JS](https://dev.office.com/fabric-js)

本文逐步展示了使用 Fabric JS 的基础知识。  

## <a name="add-the-fabric-cdn-references"></a>添加 Fabric CDN 引用
若要从 CDN 引用 Fabric，请在页面中添加以下 HTML 代码。

    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css">
    <script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>

## <a name="use-fabric-js-ux-components"></a>使用 Fabric JS 用户体验组件

Fabric JS 提供了多个可在外接程序中使用的用户体验组件，如按钮或复选框。下面列出了我们建议用于外接程序的 Fabric JS 用户体验组件。若要在外接程序中使用其中一个 Fabric 组件，请单击相应的 Fabric 文档链接，然后按**使用此组件**中的说明操作。 

- [痕迹导航](https://dev.office.com/fabric-js/Components/Breadcrumb/Breadcrumb.html)
- [按钮](https://dev.office.com/fabric-js/Components/Button/Button.html)（考虑在外接程序中使用小型按钮变体。将 16px 填充添加到小型按钮，以确保触摸设备上 40px 的最小触摸目标。）
- [复选框](https://dev.office.com/fabric-js/Components/CheckBox/CheckBox.html)
- [ChoiceFieldGroup](https://dev.office.com/fabric-js/Components/ChoiceFieldGroup/ChoiceFieldGroup.html)
- [日期选取器](https://dev.office.com/fabric-js/Components/DatePicker/DatePicker.html)（有关如何在外接程序中实现日期选取器的示例，请参阅 [Excel 销售额跟踪程序](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)代码示例。）
- [下拉列表](https://dev.office.com/fabric-js/Components/Dropdown/Dropdown.html)
- [标签](https://dev.office.com/fabric-js/Components/Label/Label.html)
- [链接](https://dev.office.com/fabric-js/Components/Link/Link.html)
- [列表](https://dev.office.com/fabric-js/Components/List/List.html)（请考虑在 CSS 中更改组件的默认样式。）
- [MessageBanner](https://dev.office.com/fabric-js/Components/MessageBanner/MessageBanner.html)
- [MessageBar](https://dev.office.com/fabric-js/Components/MessageBar/MessageBar.html)
- [覆盖](https://dev.office.com/fabric-js/Components/Overlay/Overlay.html)
- [面板](https://dev.office.com/fabric-js/Components/Panel/Panel.html)
- [透视](https://dev.office.com/fabric-js/Components/Pivot/Pivot.html)
- [ProgressIndicator](https://dev.office.com/fabric-js/Components/ProgressIndicator/ProgressIndicator.html)
- [搜索框](https://dev.office.com/fabric-js/Components/SearchBox/SearchBox.html)
- [缓冲图标](https://dev.office.com/fabric-js/Components/Spinner/Spinner.html)
- [表](https://dev.office.com/fabric-js/Components/Table/Table.html)
- [TextField](https://dev.office.com/fabric-js/Components/TextField/TextField.html)
- [开关](https://dev.office.com/fabric-js/Components/Toggle/Toggle.html)
   
## <a name="updating-your-add-in-to-use-fabric-js"></a>将外接程序更新为使用 Fabric JS
如果你一直使用的是旧版 Office UI Fabric，并且想迁移到 Fabric JS，请务必了解新组件，并在外接程序中合并和测试新组件。请注意以下几点，它们有助于你进行更新规划：

- 使用 Fabric JS 时，组件初始化更加简单。对于旧版 Fabric，需要先在外接程序项目中添加 Fabric 组件的 JavaScript 文件（包括对该文件的 `<Script>` 引用），然后初始化组件。在 Fabric JS 中，不再需要添加 Fabric 组件的 JavaScript 文件及关联的 `<Script>` 引用。只需初始化 Fabric 组件即可。   
- 多个组件现在提供可控制用户体验组件行为的函数。例如，复选框控件具有 `toggle` 函数，可以在选中和取消选中状态之间进行切换。 
- 更新了某些图标类名和样式。
- 最明显的变化是在多个组件中使用 `<label>` 元素。`<label>` 元素控制组件样式。可能需要更新用户体验代码，才能使用 `<label>` 元素。例如，更改 Fabric JS 复选框上 `<input>` 元素的 checked 属性值对复选框不会产生任何影响。请改用 `check`、`unCheck` 或 `toggle` 函数。   

## <a name="next-steps"></a>后续步骤
若要获得端到端代码示例以了解如何使用 Fabric JS，我们已经为你准备好了。请参阅以下资源：

- [Excel 销售额跟踪程序](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 

## <a name="related-resources"></a>相关资源
若要获得有关旧版 Fabric 的代码示例或文档，请参阅以下资源：

- [用户体验设计模式（使用 Fabric 2.6.1）](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Office 外接程序 Fabric UI 示例（使用 Fabric 1.0）](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [在 Office 外接程序中使用 Fabric 2.6.1](https://dev.office.com/docs/add-ins/design/ui-elements/using-office-ui-fabric)
 


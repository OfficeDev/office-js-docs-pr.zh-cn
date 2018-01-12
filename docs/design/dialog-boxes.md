# <a name="dialog-boxes-in-office-add-ins"></a>Office 外接程序中的对话框
 
对话框是浮动在活动的 Office 应用程序窗口之上的界面。你可以使用对话框为无法直接在任务窗格中打开的任务（例如登录页）提供额外的屏幕空间，或请求确认用户进行的操作，或显示如果局限在任务窗格中可能过小的视频。

**示例：对话框**

![显示对话框的典型布局的示例图像](../../images/overview_withApp_dialog.png)

### <a name="best-practices"></a>最佳做法

|**允许事项**|**禁止事项**|
|:-----|:--------|
|<ul><li>包括包含外接程序名称以及当前任务的描述性标题。</li></ul>|<ul><li>请勿在标题中追加公司名称。</li></ul>|
||<ul><li>除非方案需要，否则请勿打开对话框。</li></ul>|

## <a name="implementation"></a>实现

对于实现对话框的示例，请参阅 GitHub 中的 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)

## <a name="additional-resources"></a>其他资源

- [UX 模式示例](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
- [GitHub 开发资源](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Dialog 对象](https://dev.office.com/reference/add-ins/shared/officeui.dialog)



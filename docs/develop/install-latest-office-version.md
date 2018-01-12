# <a name="install-the-latest-version-of-office-2016"></a>安装最新版本 Office 2016

新的开发人员功能，包括那些仍处于预览状态的功能，会先发送给选择获取最新 Office 版本的订阅者。选择使用 Office 2016 的最新版本： 

- 如果您是 Office 365 家庭版、个人版或大专院校版的订阅者，请参阅[成为 Office Insider](https://products.office.com/en-us/office-insider)。
- 如果你是 Office 365 商业版客户，请参阅 [为 Office 365 商业版客户安装首次发布](https://support.office.com/en-us/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead?ui=en-US&rs=en-US&ad=US)。
- 如果在 Mac 上运行 Office 2016：
    - 启动 Office 2016 for Mac 程序。
    - 选择“帮助”菜单上的“**检查更新**”。
    - 在“Microsoft AutoUpdate”框中，选中框以加入 Office 预览体验成员计划。 

获取最新版本： 

1. 下载 [Office 2016 部署工具](https://www.microsoft.com/en-us/download/details.aspx?id=49117)。 
2. 运行该工具。这会提取以下两个文件：Setup.exe 和 configuration.xml。
3. 将 configuration.xml 文件替换为[首次发布配置文件](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)。
4. 以管理员身份运行以下命令：`setup.exe /configure configuration.xml` 

>**注意：**命令可能需要运行很长时间，而不指示进度。

在安装进程完成后，你已安装最新的 Office 2016 应用程序。要验证你是否拥有最新版本，请从任何 Office 应用程序转到“**文件**” > “**帐户**”。在“Office 更新”下，你将看到版本号上面的 (Office Insiders) 标签。

![显示产品信息的屏幕截图（带有 Office Insiders 标签）](../../images/officeinsider.PNG)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Office JavaScript API 要求集对应的最低 Office 内部版本

若要了解 API 要求集对应的各个平台的最低产品内部版本，请参阅以下资源：

- [Word JavaScript API 要求集](../../reference/requirement-sets/word-api-requirement-sets.md)
- [Excel JavaScript API 要求集](../../reference/requirement-sets/excel-api-requirement-sets.md)
- [OneNote JavaScript API 要求集](../../reference/requirement-sets/onenote-api-requirement-sets.md)
- [Dialog API 要求集](../../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Office 通用 API 要求集](../../reference/requirement-sets/office-add-in-requirement-sets.md)

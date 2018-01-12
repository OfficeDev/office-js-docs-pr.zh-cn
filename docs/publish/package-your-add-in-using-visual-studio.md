# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>使用 Visual Studio 打包加载项以准备发布

Office 加载项包包含 XML [清单文件](../overview/add-in-manifests.md)，它可用于发布加载项。 必须单独发布项目的 Web 应用程序文件。 本文介绍如何使用 Visual Studio 2015 部署 Web 项目并打包加载项。

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a>使用 Visual Studio 2015 部署 Web 项目

完成以下步骤以使用 Visual Studio 2015 部署 Web 项目。

1. 在“解决方案资源管理器”****中，打开加载项项目的快捷菜单，然后选择“发布”****。
    
    将显示“**发布外接程序**”页。
    
2. 在“当前配置文件”****下拉列表中，选择一个配置文件或选择“新建…”****以创建一个新配置文件。
    
     >**注意：**发布配置文件指定要部署到的服务器、登录服务器所需的凭据、要部署的数据库和其他部署选项。

    如果选择“新建...”****，将会显示“创建发布配置文件”****向导。 可以使用此向导从托管提供程序（如 Microsoft Azure）的网站导入发布配置文件，或创建新配置文件并添加你的服务器、凭据以及下一过程中的其他设置。
    
    有关导入发布配置文件或创建新发布配置文件的详细信息，请参阅 [创建发布配置文件](http://msdn.microsoft.com/zh-CN/library/dd465337.aspx#creating_a_profile)。
    
3. 在“**发布外接程序**”页中，选择“**部署 Web 项目**”链接。
    
    将显示“**发布 Web**”对话框。 有关使用此向导的详细信息，请参阅 [操作说明：使用 Visual Studio 中的一键式发布来部署 Web 项目](http://msdn.microsoft.com/zh-CN/library/dd465337.aspx)。
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a>使用 Visual Studio 2015 打包加载项

完成以下步骤以使用 Visual Studio 2015 打包加载项。

1. 在“发布加载项”****页中，选择“打包加载项”****链接。
    
    将显示“发布 Office 和 SharePoint 加载项”****向导。
    
2. 在“你的网站托管在哪里?”****下拉列表中，选择或输入托管加载项内容文件的网站的 URL，然后选择“完成”****。
    
    必须指定以 HTTPS 前缀开头的地址来完成此向导。 虽然通常建议使用网站的 HTTPS 终结点，但如果不打算将加载项发布到 Office 应用商店，则不需要这样做。 如果想要使用网站的 HTTP 终结点，则可以在创建包后使用文本编辑器打开 XML 清单文件，并用网站的 HTTP 前缀替换 HTTPS 前缀。 有关详细信息，请参阅[为什么我的应用和加载项必须采用 SSL 保护？](http://msdn.microsoft.com/zh-CN/library/jj591603#bk_q7)
    
     >**注意：**Azure 网站自动提供 HTTPS 终结点。

    Visual Studio 生成发布加载项所需的文件，然后打开发布输出文件夹。 
    
如果计划将加载项提交到 Office 应用商店，可以选择“执行验证检查”****链接以确定将阻止加载项被接受的任何问题。 应先解决所有问题，再将加载项提交到应用商店。

现在，可以将 XML 清单上传到适当位置，以[发布加载项](../publish/publish.md)。 可以在 `app.publish` 文件夹的 `OfficeAppManifests` 中找到 XML 清单： 例如：

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="additional-resources"></a>其他资源



- [发布 Office 外接程序](../publish/publish.md)
    
- [将 Office 与 SharePoint 外接程序和 Office 365 Web 应用提交到 Office 应用商店](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    

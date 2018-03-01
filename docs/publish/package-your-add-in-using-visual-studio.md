---
title: 使用 Visual Studio 打包加载项以准备发布
description: ''
ms.date: 01/25/2018
---


# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>使用 Visual Studio 打包加载项以准备发布

Office 加载项包包含 XML [清单文件](../develop/add-in-manifests.md)，它可用于发布加载项。 必须单独发布项目的 Web 应用程序文件。 本文介绍如何使用 Visual Studio 2015 部署 Web 项目并打包加载项。

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a>使用 Visual Studio 2015 部署 Web 项目

完成以下步骤以使用 Visual Studio 2015 部署 Web 项目。

1. 在“解决方案资源管理器”****中，打开加载项项目的快捷菜单，然后选择“发布”****。
    
    将显示“**发布外接程序**”页。
    
2. 选择“当前配置文件”****下拉列表中的配置文件，或选择“新建…”****新建配置文件。
    
    > [!NOTE]
    > 发布配置文件指定要部署到的服务器、登录服务器所需的凭据、要部署的数据库和其他部署选项。

    如果你选择“**新建...**”，将会显示“**创建发布配置文件**”向导。可以使用此向导从托管提供程序（如 Microsoft Azure）的网站导入发布配置文件，或创建新配置文件并添加你的服务器、凭据以及下一过程中的其他设置。
    
    有关导入发布配置文件或创建新发布配置文件的详细信息，请参阅 [创建发布配置文件](http://msdn.microsoft.com/zh-cn/library/dd465337.aspx#creating_a_profile)。
    
3. 在“**发布外接程序**”页中，选择“**部署 Web 项目**”链接。
    
    The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/zh-cn/library/dd465337.aspx).
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a>使用 Visual Studio 2015 打包加载项的具体步骤

完成以下步骤以使用 Visual Studio 2015 打包加载项。

1. 在“发布加载项”****页中，选择“打包加载项”****链接。
    
    此时，“发布 Office 和 SharePoint 加载项”****向导显示。
    
2. 在“网站托管在哪里?”****下拉列表中，选择或输入托管加载项内容文件的网站的 HTTPS URL，再选择“完成”****。 
    
    必须指定以 HTTPS 前缀开头的 URL，才能完成此向导。若要使用网站的 HTTP 终结点，可以在创建包后使用文本编辑器打开 XML 清单文件，并将网站的 HTTPS 前缀替换为 HTTP 前缀。 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure 网站自动提供 HTTPS 终结点。

    此时，Visual Studio 生成发布加载项所需的文件，并打开发布输出文件夹。 
    
如果计划将加载项提交到 AppSource，可以选择“执行验证检查”****链接，以发现任何可能会导致加载项遭拒的问题。 应先解决所有问题，再将加载项提交到 Microsoft Store。

现在，可以将 XML 清单上传到适当位置，以[发布加载项](../publish/publish.md)。XML 清单位于 `app.publish` 文件夹的 `OfficeAppManifests` 中。例如：

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a>另请参阅

- [发布 Office 加载项](../publish/publish.md)
- 
  [将解决方案提交到 AppSource 和 Office 应用商店](https://docs.microsoft.com/zh-cn/office/dev/store/submit-to-the-office-store)
    

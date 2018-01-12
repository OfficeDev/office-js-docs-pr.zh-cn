# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a>将任务窗格和内容加载项发布到 SharePoint 目录

加载项目录是 SharePoint Web 应用程序或 SharePoint Online 租户中的专用网站集，用于托管 Office 和 SharePoint 加载项的文档库。为使组织内的用户可访问 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的加载项目录中。 管理员将外接程序目录注册为受信任的目录时，用户可从 Office 客户端应用程序中的插入 UI 中插入外接程序。

**重要说明**： 

- SharePoint 上的加载项目录不支持在[加载项清单](../overview/add-in-manifests.md)的 `VersionOverrides` 节点中实现的加载项功能，如加载项命令。

- 如果面向的是云或混合环境，建议通过 [Office 365 管理中心使用集中部署](publish/centralized-deployment.md)来发布加载项。

- SharePoint 目录不支持 Office 2016 for Mac。 若要向 Mac 客户端部署 Office 加载项，必须将其提交到 [Office 应用商店](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)。   

## <a name="set-up-an-add-in-catalog"></a>设置加载项目录

完成以下部分之一中的步骤，以在 SharePoint 或 Office 365 上设置加载项目录。

### <a name="to-set-up-an-add-in-catalog-on-sharepoint"></a>在 SharePoint 上设置加载项目录

1. 浏览到“**管理中心网站**”（“**开始**” > “**所有程序**” > “**Microsoft SharePoint 2013 产品**” > “**SharePoint 2013 管理中心**”）。
    
2. 在左侧的任务窗格中，选择“**外接程序**”。
    
3. 在“**外接程序**”页面的“**外接程序管理**”下，选择“**管理外接程序目录**”。
    
4. 在“**管理外接程序目录**”页上，确保在“**Web 应用程序选择器**”中选择了正确的 Web 应用程序。
    
5. 选择“**查看网站设置**”。
    
6. 在“**网站设置**”页上选择“**网站集管理员**”以指定网站集管理员，然后选择“**确定**”。
    
7. 若要向用户授予网站权限，请选择“**网站权限**”，然后选择“**授予权限**”。
    
8. 在“**共享‘应用程序目录网站’**”对话框中，指定一个或多个网站用户，为他们设置相应的权限，选择性地设置其他选项，然后选择“**共享**”。
    
9. 若要向 Office 加载项加载项目录添加加载项，请选择“Office 加载项”****。

### <a name="to-set-up-an-add-in-catalog-on-office-365"></a>在 Office 365 上设置加载项目录

1. 在“Office 365 管理中心”页上，选择“**管理**”，然后选择“**SharePoint**”。
    
2. 在左侧的任务窗格中，选择“**外接程序**”。
    
3. 在“**外接程序**”页上，选择“**外接程序目录**”。
    
4. 在“**外接程序目录网站**”页上，选择“**确定**”以接受默认选项，并新建外接程序目录网站。
    
5. 在“**创建外接程序目录网站集**”页上，指定外接程序目录网站的标题。
    
6. 指定网站地址。
    
7. 将“**存储配额**”设置为可能的最低值（当前为 110）。 你将仅在该网站集上安装外接程序包，它们非常小。
    
8. 将“**服务器资源配额**”设置为 0（零）。 （服务器资源配额与限制性能不佳的沙盒解决方案有关，但你不会在外接程序目录网站上安装任何沙盒解决方案。）
    
9. 选择“确定”****。
    
10. 若要将加载项添加到加载项目录网站，请浏览至刚刚创建的网站。 在左侧导航窗格中，选择“Office 加载项”****，然后选择“新加载项”****以上传 Office 加载项清单文件。

## <a name="publish-an-add-in-to-an-add-in-catalog"></a>将加载项发布到加载项目录

若要将加载项发布到加载项目录，请完成以下步骤。

1. 浏览至加载项目录：

    1- 打开 SharePoint 管理中心主页。
    
    2- 选择“**外接程序**”。
    
    3- 选择“**管理外接程序目录**”。
    
    4- 选择提供的链接，然后选择左侧导航栏上的“**Office 外接程序**”。
    
2. 选择“**单击以添加新项目**”链接。
    
3. 选择“**浏览**”，然后指定要上载的 [清单](../../docs/overview/add-in-manifests.md)。
    
    此目录中的内容和任务窗格外接程序现在可从“**Office 外接程序**”对话框提供。 若要访问这些加载项，请在“插入”****选项卡上选择“我的加载项”****，然后选择“我的组织”****。

## <a name="end-user-experience-with-the-add-in-catalog"></a>加载项目录最终用户体验

最终用户可以通过完成以下步骤来访问 Office 应用程序中的加载项目录：

1. 在 Office 应用程序中，转到“文件”**** > “选项”****“信任中心” > **** > 信任中心设置**** > “受信任的加载项目录”****。
    
2. 指定加载项目录的_父级 SharePoint 网站集_的 URL。 例如，如果 Office 加载项目录的 URL 是：
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    仅指定父网站集的 URL：
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. 关闭并重新打开 Office 应用程序。 加载项目录会出现在“Office 加载项”****对话框中。

或者，管理员可以使用组策略在 SharePoint 上指定 Office 加载项目录。 有关详细信息，请参阅 TechNet 上的[使用组策略管理用户安装和使用 Office 加载项的方式](https://technet.microsoft.com/zh-CN/library/jj219429.aspx#BKMK_GP)一节。


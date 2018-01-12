
# <a name="host-an-office-add-in-on-microsoft-azure"></a>在 Microsoft Azure 上托管 Office 外接程序

最简单的 Office 外接程序由 XML 清单文件和 HTML 页构成。XML 清单文件描述了外接程序的特性，例如它的名称、可以运行它的 Office 客户端应用程序以及外接程序 HTML 页的 URL。HTML 页包含在一个 Web 应用中，用户在 Office 客户端应用程序中安装和运行外接程序时将与此 Web 应用进行交互。可以将 Office 外接程序的 Web 应用托管在任意 Web 托管平台（包括 Azure）上。

本文介绍了如何将外接程序 Web 应用部署到 Azure 并[旁加载外接程序](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)以在 Office 客户端应用程序中进行测试。

## <a name="prerequisites"></a>先决条件 

1. 安装 [Visual Studio 2017](https://www.visualstudio.com/downloads) 并选择包含 **Azure 开发**工作负载。

    >**注意：**如果之前已安装 Visual Studio 2017，请[使用 Visual Studio 安装程序](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio)，以确保安装 **Azure 开发**工作负载。 

2. 安装 Office 2016。 
    
     >**注意：**如果尚未安装 Office 2016，可以[注册 1 个月免费试用版](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786)。

3.  获取 Azure 订阅。
    
     >**注意：**如果还没有 Azure 订阅，可以[通过 MSDN 订阅获取 Azure 订阅](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/)或[注册免费试用版](https://azure.microsoft.com/en-us/pricing/free-trial)。 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a>步骤 1：创建一个共享文件夹，用于托管外接程序 XML 清单文件

1. 打开开发计算机的文件资源管理器。
    
2. 右键单击 C:\ 驱动器，然后选择“新建”**** > “文件夹”****。
    
3. 将新文件夹命名为 AddinManifests。
    
4. 右键单击 AddinManifests 文件夹，然后选择“共享”**** > “特定用户”****。
    
5. 在“文件共享”****中，选择下拉箭头，然后选择“所有人”**** > “添加”**** > “共享”****。
    
> **注意：**在本演练中，将本地文件共享作为存储外接程序 XML 清单文件的受信任目录。在实际方案中，可以改为选择[将 XML 清单文件部署到 SharePoint 目录](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)或[将外接程序发布到 Office 应用商店](https://dev.office.com/officestore/docs/submit-to-the-office-store)。

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a>第 2 步：将文件共享添加到受信任的外接程序目录

1.  启动 Word 2016 并创建文档。

    >**注意：**尽管本示例使用的是 Word 2016，但你可以使用任何支持 Office 外接程序的 Office 应用程序，如 Excel、Outlook、PowerPoint 或 Project 2016。
    
2.  选择“文件”**** > “选项”****。
    
3.  在“Word 选项”****对话框中，选择“信任中心”****，然后选择“信任中心设置”****。 
    
4.  在“信任中心”****对话框中，选择“受信任的外接程序目录”****。输入之前创建的文件共享的通用命名约定 (UNC) 路径，作为**目录 URL**（例如，\\\YourMachineName\AddinManifests）。然后选择“添加目录”****。 
    
5. 选中“在菜单中显示”****复选框。 

    >**注意：**将外接程序 XML 清单文件存储在已指定为受信任的 Web 外接程序目录的共享中时，如果用户导航到功能区中的“插入”****选项卡并选择“我的外接程序”****，则外接程序显示在“Office 外接程序”****对话框中的“共享文件夹”****下。

6. 关闭 Word 2016。

## <a name="step-3-create-a-web-app-in-azure"></a>第 3 步：在 Azure 中创建 Web 应用

使用 [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) 或使用 [Azure 门户](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal)在 Azure 中创建空的 Web 应用。

### <a name="using-visual-studio-2017"></a>使用 Visual Studio 2017

若要使用 Visual Studio 2017 创建 Web 应用，请完成以下步骤。

1. 在 Visual Studio 的“视图”****菜单中，选择“服务器资源管理器”****。右键单击“Azure”**** 并选择“连接到 Microsoft Azure 订阅”****。请按说明连接到 Azure 订阅。
    
2. 在 Visual Studio 的“服务器资源管理器”****中，展开“Azure”****，右键单击“应用服务”****，然后选择“创建新的应用服务”****。
    
3. 在“创建应用服务”****对话框中，提供此信息：
    
      - 为站点输入唯一的“Web 应用名称”****。Azure 验证站点名称在整个 azurewebsites.net 域中是否是唯一的。

      - 选择要用于创建此站点的“订阅”****。

      - 为站点选择“资源组”****。如果创建新组，还需为新组命名。
    
      - 选择要用于创建此站点的“应用服务计划”****。如果创建新计划，还需为新计划命名。
       
      - 选择“创建”****。

    新的 Web 应用将在“服务器资源管理器”****中的“Azure”**** >> “应用服务”****>>“选择的资源组”下显示。
    
4. 右键单击新的 Web 应用，然后选择“在浏览器中查看”****。随即打开浏览器，并显示包含“应用服务应用已创建”消息的网页。
    
5. 在浏览器地址栏中，更改 Web 应用的 URL，以便其使用 HTTPS，并按 **Enter** 确认已启用 HTTPS 协议。Office 外接程序模型要求外接程序使用 HTTPS 协议。
    
### <a name="using-the-azure-portal"></a>使用 Azure 门户

若要使用 Azure 门户创建 Web 应用，请完成以下步骤。

1. 使用 Azure 凭据登录到 [Azure 门户](https://portal.azure.com/)。
    
2. 选择“新建”**** > “Web + 移动”**** > “Web 应用”****。 

3. 在“Web 应用创建”****对话框中，提供此信息：
    
      - 为站点输入唯一的“应用名称”****。Azure 验证站点名称在整个 azureweb apps.net 域中是否是唯一的。

      - 选择要用于创建此站点的“订阅”****。

      - 为站点选择“资源组”****。如果创建新组，还需为新组命名。

      - 为站点选择“操作系统”****。
    
      - 选择要用于创建此站点的“应用服务计划”****。如果创建新计划，还需为新计划命名。
       
      - 选择“创建”****。

4. 选择“通知”****（Azure 门户顶部边缘的钟形图标），然后选择“部署成功”****通知，以打开 Azure 门户中的站点“概述”****页。

    >**注意：**站点部署完成后，通知从“正在部署”****更改为“部署成功”****。

5. 在 Azure 门户的站点“概述”****页的“基本信息”****部分中，选择“URL”****下显示的 URL。随即打开浏览器，并显示包含“应用服务应用已创建”消息的网页。 
    
6. 在浏览器地址栏中，更改 Web 应用的 URL，以便其使用 HTTPS，并按 **Enter** 确认已启用 HTTPS 协议。Office 外接程序模型要求外接程序使用 HTTPS 协议。    

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a>步骤 4：在 Visual Studio 中创建 Office 外接程序

1. 以管理员身份启动 Visual Studio。
    
2. 选择“文件”**** > “新建”**** > “项目”****。
    
3. 在“模板”****下，展开“Visual C#”****（或“Visual Basic”****），展开“Office/SharePoint”****，然后选择“外接程序”****。
    
4. 选择“Word Web 外接程序”****，然后选择“确定”****以接受默认设置。
       
Visual Studio 将创建基本的 Word 外接程序，你可以按原样发布，无需对其 Web 项目进行任何更改。

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a>第 5 步：将 Office 外接程序 Web 应用发布到 Azure

1. 在 Visual Studio 中打开外接程序项目后，展开“解决方案资源管理器”****中的解决方案节点，以便可以查看解决方案的两个项目。
    
2. 右键单击 Web 项目，然后选择“发布”****。Web 项目包含 Office 外接程序 Web 应用文件，因此，这是你可以发布到 Azure 的项目。
    
3. 在“发布”****选项卡上：

      - 选择“Microsoft Azure 应用服务”****。
      
      - 选择“选择现有”****。

      - 选择“发布”****。 

6. 在“应用服务”****对话框中，找到并选择在[步骤 3：在 Azure 中创建 Web 应用](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure)中创建的 Web 应用，然后选择“确定”****。 

    Visual Studio 会将 Office 外接程序的 Web 项目发布到 Azure Web 应用。Visual Studio 完成发布 Web 项目后，浏览器将打开并显示网页，其中显示“应用服务应用已创建”文本。这是 Web 应用当前的默认页。

7. 若要查看外接程序的网页，请更改此 URL，使其使用 HTTPS，并指定外接程序的 HTML 页的路径（例如：https://YourDomain.azurewebsites.net/Home.html）。这可确认你的外接程序的 Web 应用现在托管于 Azure 上。复制根 URL（例如：https://YourDomain.azurewebsites.net）；本文后续部分中编辑外接程序清单文件时将需要此 URL。
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a>步骤 6：编辑并部署外接程序 XML 清单文件

1. 在示例 Office 外接程序在“解决方案资源管理器”****中打开的 Visual Studio 中，展开该解决方案以显示两个项目。
    
2. 展开 Office 外接程序项目（例如 WordWebAddIn），右键单击清单文件夹，然后选择“打开”****。随即打开外接程序 XML 清单文件。
    
3. 在 XML 清单文件中，找到所有的“~remoteAppUr”实例，并将其全部替换为 Azure 上的外接程序 Web 应用的根 URL。这就是之前在将外接程序 Web 应用发布到 Azure 后复制的 URL（例如：https://YourDomain.azurewebsites.net）。 
    
4. 选择“文件”****，然后选择“全部保存”****。关闭外接程序 XML 清单文件。
    
5. 返回到“解决方案资源管理器”****，右键单击清单文件夹并选择“在文件资源管理器中打开文件夹”****。
    
6. 复制外接程序 XML 清单文件（例如 WordWebAddIn.xml）。 
    
7. 浏览到在[步骤 1：创建共享文件夹](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file)中创建的网络文件共享，并将清单文件粘贴到此文件夹。

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a>步骤 7：在 Office 客户端应用程序中插入并运行加载项

1. 启动 Word 2016 并创建文档。
    
2. 在功能区中选择“插入”**** > “我的外接程序”****。 
    
3. 在“Office 外接程序”****对话框中，选择“共享文件夹”****。Word 扫描已列为受信任的外接程序目录（在[步骤 2：将文件共享添加到受信任的外接程序目录](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)）的文件夹，并在对话框中显示外接程序。应该会看到示例外接程序的图标。
    
4. 选择你的外接程序的图标，然后选择“添加”****。外接程序的“显示任务窗格”****按钮将添加到功能区。 

5. 在“主页”****选项卡的功能区上，选择“显示任务窗格”****按钮。外接程序在当前文档右侧的任务窗格中打开。
    
6. 通过选中文档中的一些文本并在任务窗格中选择“高亮颜色!”****按钮来验证外接程序是否可正常工作。 

## <a name="additional-resources"></a>其他资源

- [发布 Office 外接程序](../publish/publish.md)
    
- [使用 Visual Studio 打包外接程序以准备发布](../publish/package-your-add-in-using-visual-studio.md)
    


# <a name="get-started-with-labsjs-for-office-mix"></a>适用于 Office Mix 的 LabsJS 入门



LabsJS 内容公开了您可用于开发交互式实验室的 API (labs.js)、示例、文档和相关文件，将其集成到 Office Mix，然后将其呈现在 Microsoft PowerPoint 中。这些实验室实际上是您使用 HTML5 和 labs.js JavaScript 库创建的 Office 外接程序。

## <a name="labsjs-content"></a>LabsJS 内容

LabsJS 提供了创建和发布你自己的 Office Mix 实验室所需的文档、示例实验室和文件。


**所需文件**


|**文件**|**说明**|
|:-----|:-----|
|labs-1.0.4.js|用于开发 Office Mix 实验室的 LabsJS JavaScript API。此文件必须包含在你的项目中，以允许其与 Office Mix 集成。此文件还会承载在内容交付网络 (CDN) <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code> 上。当你发布应用程序时，必须链接到 CDN 上的文件。|
|labs-1.0.4.d.ts|labs.js 的 typeScript 定义文件。这样就可以轻松地将 TypeScript 代码与 labs.js 集成。定义文件还提供了 labs.js 中包含的所有组件的全面概述。您可以从 [http://www.typescriptlang.org/](http://www.typescriptlang.org/) 下载 TypeScript。定义文件针对 TypeScript 版本 0.9.1.1 生成。|
|历史记录|labs.js 库各个版本的发布历史记录。|
|Labshost.html|一个网页，允许您在 PowerPoint 上下文以外针对 Office Mix 查看和调试实验室。要使用此页面，请在主输入框中键入您的 URL，然后它将加载在框中。在 PowerPoint 或 Office Mix 课程播放器中运行时在 API 和 Office Mix 之间交换的数据将显示在右侧的输入框中。数据也可以预先设定种子。请注意，"示例"部分中的示例实验室显示了在主机上下文中运行的现有 Office Mix 外接程序。|
|SampleManifest.xml|一个示例 Office 外接程序 清单，用作创建您自己的应用程序清单的模板。|
|Simplelab.html|使用 labs.js 创建的示例实验室。允许选择和插入网页，该网页将跟踪进行查看的用户。|
|Simplelab.ts|用于创建 simplelab 示例的 TypeScript 文件。|
|Simplelab.js|Simplelab 示例的 JavaScript 版本。此文件和 simplelab.ts 显示了 LabsJS API 的用途。|

## <a name="set-up-your-development-environment"></a>设置开发环境

labs.js 库充当除 office.js 库（Office 外接程序 的 API）以外的抽象层，因此您使用 labs.js 库创建的实验室实际上是 Office 外接程序。为了能够使用 labs.js 库并在 Office Mix 内部运行这些实验室，您必须首先将自己设置为 Office 外接程序开发人员。


### <a name="register-for-an-office-365-developer-site"></a>注册 Office 365 开发人员网站

第一步是注册 Office 365 开发人员网站。这样您就可以在将实验室提交到 Office 应用商店之前对其进行托管和测试。此网站允许您在实时环境中将外接程序发布到 Office Mix 并进行测试。

有关详细信息，请参阅[在 Office 365 上为 SharePoint 外接程序设置开发环境](http://msdn.microsoft.com/library/b22ce52a-ae9e-4831-9b68-c9210af6dc54%28Office.15%29.aspx) 


### <a name="set-up-an-app-catalog-on-sharepoint-online"></a>在 SharePoint Online 上设置应用程序目录

创建和设置开发人员网站之后，您可在 SharePoint Online 上设置外接程序。有关详细信息，请参阅 [在 Office 365 上设置外接程序目录](../../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。

对于 Office Mix，您使用外接程序目录，以便可以在将实验室提交到应用商店之前将预生产外接程序插入到课程中并执行端到端测试。


## <a name="create-your-lab"></a>创建实验室

要创建第一个实验室，请按照 [演练](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)中的步骤执行操作，其中介绍了如何创建简单的正/误判断测验。请参阅 [演练：为 Office Mix 创建第一个实验室](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)。


## <a name="publish-your-lab"></a>发布实验室

创建实验室之后，您可以将其发布并提交到应用商店。


### <a name="create-and-upload-your-application-manifest"></a>创建并上载应用程序清单

应用程序清单是描述您的 LabJS 实验室的 XML 文档。它提供了对托管实验室的 URL 的引用，并包含有关实验室的详细信息，包括显示名称、描述、图标、大小等。

我们包含了示例清单"SampleManifest.xml"。有关清单架构的详细信息以及架构定义链接，请参阅 [Office 外接程序 XML 清单](../../../docs/overview/add-in-manifests.md)。

若要将清单上载到 SharePoint 网站，首先导航到应用程序目录，通常可以通过此 URL <code>https://\<your site\>/sites/AppCatalog</code> 查找到该目录。然后，选择“**新的应用**”按钮，然后按照步骤上载应用程序清单。


### <a name="update-your-powerpoint-2013-catalog"></a>更新 PowerPoint 2013 目录

接下来，更新您的 PowerPoint 2013 目录，然后使用您的开发人员帐户登录。

从更新 PowerPoint 2013 目录开始启动 PowerPoint 2013 并导航菜单路径“**文件 > 选项 > 信任中心 > 信任中心设置 > 受信任的应用程序目录**”。由此处向应用程序目录添加引用，然后选择“**确认**”。PowerPoint 2013 将要求你进行注销以使更改生效。注销。

最后，使用开发人员帐户重新登录。在 PowerPoint 2013 右上角选择您的登录名称并使用您的开发人员帐户登录。现在您可以插入外接程序。


### <a name="insert-publish-and-view-your-app"></a>插入、发布和查看应用程序

若要将外接程序插入此目录，请选择“**插入**”功能区，然后在“**应用程序**”区域中选择“**存储**”。选择“**我的组织**”，然后将在外接程序目录中看到此外接程。选择外接程序，选择“**插入**”，则外接程序（实验室）会插入到 PowerPoint 2013 文档中。

现在，你可以利用所有可用的 Office Mix 功能在新实验室中发布课程。


 >**重要信息**：要查看应用程序，你必须使用查看课程的相同浏览器登录到你的 SharePoint 目录。SharePoint 目录仅允许经过身份验证的用户访问，因此要查看应用程序，你需要先登录。 


### <a name="submit-your-lab-to-the-office-store"></a>将实验室提交到 Office 应用商店

要将实验室提交到 Office 应用商店，请参阅[发布 Office 外接程序](../../publish/publish.md)


## <a name="additional-resources"></a>其他资源



- [Office Mix 外接程序](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Office 外接程序](../../../docs/overview/office-add-ins.md)
    
- [创建第一个 Office Mix 实验室](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)
    

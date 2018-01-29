
# <a name="office-add-ins-development-lifecycle"></a>Office 外接程序开发生命周期

> [!NOTE]
> 生成外接程序时，如果计划将外接程序[发布](../publish/publish.md)到 Office 应用商店，请务必遵循 [Office 应用商店验证策略](https://msdn.microsoft.com/zh-cn/library/jj220035.aspx)。例如，外接程序必须适用于支持你定义的方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://msdn.microsoft.com/zh-cn/library/jj220035.aspx#Anchor_3)以及 [Office 外接程序主机和可用性](https://dev.office.com/add-in-availability)页）。

Office 外接程序的典型开发生命周期包括下列步骤：


1.  **决定外接程序的用途。**
    
    对以下问题提问：
    
      - 加载项有何作用？ 
    
      - 它如何帮助您的客户提高工作效率？
    
      - 您的加载项功能支持哪些方案？
    

    确定最重要的功能和方案，并围绕它们进行集中设计。 
    
2.  **标识外接程序的数据和数据源。**
    
    数据是否在文档、工作簿、演示文稿、项目、基于浏览器的 Access 数据库、或有关 Exchange Server 或 Exchange Online 邮箱中的项或项目中？数据是否来自外部源（如 Web 服务）？
    
3.  **确定外接程序类型和最能支持其用途的 Office 主机应用程序。**
    
    请考虑以下各项以确定方案：
    
    - 客户是否要使用该外接程序来丰富文档或基于 Access 浏览器的数据库的内容？如果是这样，你可能要考虑创建内容外接程序。 
    
    - 客户是否要在查看或撰写电子邮件或约会时使用该外接程序？能够根据当前上下文公开外接程序是否很重要？是否优先考虑使外接程序不仅在台式机上可用，而且在平板电脑或智能手机上也可用？
    
        如果答案是肯定的，请考虑创建 Outlook 外接程序。然后确定将触发外接程序的上下文（例如，撰写窗体中的用户、特定消息类型、附件存在形式、地址、任务建议，会议建议或电子邮件或约会的内容中的特定字符串模式）。请参阅 [Outlook 外接程序的激活规则](../outlook/manifests/activation-rules.md)以了解如何根据上下文激活 Outlook 外接程序。
    
    - 客户是否要使用该外接程序来增强文档的查看或创作体验？如果是这样，你可能要考虑创建任务窗格外接程序。 

    （Windows、Mac、Web、Mobile）上运行的 Office 应用程序和平台之间的某些外接程序 API 可能不同。若要查看客户端和平台的当前 API 覆盖范围，请参阅我们的 [Office 外接程序主机和平台可用性](https://dev.office.com/add-in-availability)页。  
    
4.  **为外接程序设计并实现用户体验和用户界面。**
    
    设计快速、流畅的用户体验，该体验一致、易于学习，并且主要方案只需几个步骤即可完成。根据加载项的用途，利用第三方 API 或 Web 服务。
    
    可从各种 Web 开发工具中进行选择，并使用 HTML 和 JavaScript 执行用户界面。
    
5.  **根据 Office 外接程序清单架构创建 XML 清单文件。**
    
    创建 XML 清单，以确定加载项及其要求，指定加载项使用的 HTML 以及任何 JavaScript 和 CSS 文件的位置，并根据加载项的类型指定默认大小和权限。
    
    对于 Outlook 外接程序，可以根据当前邮件或约会指定上下文，外接程序在其中相关并使 Outlook 可在 UI 中使用。您还可以决定希望外接程序支持的设备。在清单中，指定作为激活规则的上下文和受支持的设备。
    
6.  **安装和测试外接程序。**
    
    将 HTML 文件以及任何 JavaScript 和 CSS 文件放在外接程序清单文件中指定的 Web 服务器上。安装外接程序的过程取决于外接程序的类型。有关详细信息，请参阅[旁加载 Office 外接程序进行测试](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)。
    
    对于 Outlook 外接程序，将其安装在 Exchange 邮箱中，并指定 Exchange 管理中心 (EAC) 中外接程序清单文件的位置。有关详细信息，请参阅[部署和安装 Outlook 外接程序以供测试](../outlook/testing-and-tips.md)。
    
7.  **发布外接程序。**
    
    可将外接程序提交到 Office 应用商店，客户可以从这里安装外接程序。此外，可以向 SharePoint 上的专有文件夹外接程序目录或共享网络文件夹发布任务窗格和内容外接程序，并在组织的 Exchange Server 上直接部署 Outlook 外接程序。有关详细信息，请参阅[发布 Office 外接程序](../publish/publish.md)。
    
8.  **维护外接程序**
    
    如果外接程序调用 Web 服务，且在发布外接程序后对 Web 服务进行了更新，则无需重新发布外接程序。但是，如果你对提交的外接程序的任何项目或数据进行了更改（如外接程序清单、屏幕截图、图标、HTML 或 JavaScript 文件），则需重新发布外接程序。尤其是，如果已向 Office 应用商店发布了外接程序，则需重新提交外接程序，以便 Office 应用商店可以执行这些更改。重新提交外接程序时，必须附带包括新版本号的已更新的外接程序清单。还必须确保更新提交表单中的外接程序版本号以匹配新清单的版本号。对于 Outlook 外接程序，应确保 [Id](http://dev.office.com/reference/add-ins/manifest/id) 元素在外接程序清单中包含不同的 UUID。
    

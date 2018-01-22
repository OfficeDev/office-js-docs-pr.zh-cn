# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a>更新到适用于 Office 的 JavaScript API 的最新库和第 1.1 版加载项清单架构

本文介绍了如何将 Office 外接程序项目中的 JavaScript 文件（Office.js 和特定于应用程序的 .js 文件）和外接程序清单验证文件更新到版本 1.1。

## <a name="using-the-most-up-to-date-project-files"></a>使用最新的项目文件

如果您使用 Visual Studio 来开发您的外接程序，以使用适用于 Office 的 JavaScript API 的 [最新 API 成员](../../reference/what's-changed-in-the-javascript-api-for-office.md)和 [外接程序清单 v1.1 功能](../../docs/overview/add-in-manifests.md)（根据 offappmanifest-1.1.xsd 进行了验证），则您需要下载并安装 [Visual Studio 2015 和最新的 Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs)。

如果您使用文本编辑器或 Visual Studio 以外的 IDE 开发您的 外接程序，则您需要针对在 外接程序 的清单中引用的 Office.js 和架构版本，将引用更新到 CDN。

若要运行使用新的和已更新的 Office.js API 和加载项清单功能开发的加载项，您的客户必须运行 Office 2013 SP1 或更高版本的本地产品，并在适用的情况下运行 SharePoint Server 2013 SP1 和相关的服务器产品、Exchange Server 2013 Service Pack 1 (SP1) 或相当于联机托管的产品：Office 365、SharePoint Online 和 Exchange Online。

若要下载 Office、SharePoint 和 Exchange SP1 产品，请参阅以下内容：

- [Microsoft Office 2013 和相关桌面产品的所有 Service Pack 1 (SP1) 更新的列表](http://support.microsoft.com/kb/2850036)
    
- [Microsoft SharePoint Server 2013 和相关服务器产品的所有 Service Pack 1 (SP1) 更新的列表](http://support.microsoft.com/kb/2850035)
    
- [Exchange Server 2013 Service Pack 1 的说明](http://support.microsoft.com/kb/2926248)
    

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>更新使用 Visual Studio 创建的 Office 加载项项目

对于在适用于 Office 的 JavaScript API v1.1 和外接程序清单架构发布之前创建的项目，你可以使用“**NuGet 程序包管理器**”更新项目文件，然后更新外接程序的 HTML 页以进行引用。 

请注意，更新过程对于 _每个项目_ 执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，您需要重复更新过程。


### <a name="to-update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a>将项目中适用于 Office 的 JavaScript API 库文件更新到最新版本


1. 在 Visual Studio 2015 中，打开或创建新的“**Office 外接程序**”项目。
    
      - 在左侧窗格中，选择“**更新**”并完成程序包更新过程。
    
  - 转到步骤 6。
    
2. 依次选择“**工具**” > “**NuGet 包管理器**” > “**管理解决方案的 Nuget 包**”。
    
3. 在“**NuGet 程序包管理器**”中，为“**程序包源**”选择“**nuget.org**”并为“**筛选器**”选择“**可用升级**”。 并选择 Microsoft.Office.js。
    
4. 在左侧窗格中，选择“更新”****，完成包更新流程。
    
5. 在外接程序的 HTML 页的 **head** 标记中，注释掉或删除任何现有的 office.js 脚本引用，然后引用已更新的适用于 Office 的 JavaScript API 库，方法如下（：
    
```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

> **注意**：在 CDN URL 中，`office.js` 前面的 `/1/` 指定在第 1 版 Office.js 中使用最新增量版本。   


### <a name="to-update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>将项目中的清单文件更新为使用第 1.1 版架构的具体步骤

在外接程序清单文件中，更新 **OfficeApp** 元素的 **xmlns** 属性，将版本值更改为 `1.1`（除 **xmlns** 属性以外的属性保持不变）。
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> **注意：**在将外接程序清单架构的版本更新为 1.1 之后，你将需要删除 **Capabilities** 和 **Capability** 元素，并将其替换为 [Hosts](http://dev.office.com/reference/add-ins/manifest/hosts) 和 [Host](http://dev.office.com/reference/add-ins/manifest/hosts) 元素或 [Requirements 和 Requirement 元素](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>更新使用文本编辑器或其他 IDE 创建的 Office 加载项项目

对于在发布适用于 Office 的 JavaScript API v1.1 和加载项清单架构之前创建的项目，您需要将加载项的 HTML 页更新到 v1.1 的 CDN 引用库中，将您的加载项清单文件更新为使用架构 v1.1。 

更新过程对_每个项目_分别执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，你需要重复更新过程。

你不需要适用于 Office 的 JavaScript API 文件（Office.js 和特定于应用程序的.js 文件）的本地副本来开发 Office 加载项（在运行时引用 Office.js 的 CDN 会下载必要的文件），但如果你想要库文件的本地副本，你可以使用 [NuGet 命令行实用程序](http://docs.nuget.org/consume/installing-nuget)和 `Install-Package Microsoft.Office.js` 命令来下载它们。

 > **注意：**若要获取有关 v1.1 外接程序清单的 XSD（XML 架构定义）副本，请参阅 [Office 外接程序清单的架构参考 (v1.1)](../overview/add-in-manifests.md) 中列出的内容。


### <a name="to-update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a>将项目中适用于 Office 的 JavaScript API 库文件更新为使用最新版本

1. 在文本编辑器或 IDE 中打开外接程序的 HTML 页。
    
2. 在外接程序的 HTML 页的 **head** 标记中，注释掉或删除任何现有的 office.js 脚本引用，然后引用已更新的适用于 Office 的 JavaScript API 库，方法如下（：
    
```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

> **注意**：在 CDN URL 中，`office.js` 前面的 `/1/` 指定在第 1 版 Office.js 中使用最新增量版本。   

### <a name="to-update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>将项目中的清单文件更新为使用第 1.1 版架构的具体步骤

在外接程序清单文件中，更新 **OfficeApp** 元素的 **xmlns** 属性，将版本值更改为 `1.1`（除 **xmlns** 属性以外的属性保持不变）。
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> **注意：**在将外接程序清单架构的版本更新为 1.1 之后，你将需要删除 **Capabilities** 和 **Capability** 元素，并将其替换为 [Hosts](http://dev.office.com/reference/add-ins/manifest/hosts) 和 [Host](http://dev.office.com/reference/add-ins/manifest/hosts) 元素或 [Requirements 和 Requirement 元素](../../docs/overview/specify-office-hosts-and-api-requirements.md)。
    

## <a name="additional-resources"></a>其他资源

- [指定 Office 主机和 API 要求](../../docs/overview/specify-office-hosts-and-api-requirements.md)
    
- [了解适用于 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [适用于 Office 的 JavaScript API](http://dev.office.com/reference/add-ins/javascript-api-for-office)
    
- [Office 外接程序清单的架构参考 (v1.1)](../overview/add-in-manifests.md)
    

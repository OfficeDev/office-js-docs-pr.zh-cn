---
title: 更新到适用于 Office 的 JavaScript API 最新库和第 1.1 版加载项清单架构
description: 将 Office 加载项项目中使用的 JavaScript 文件（Office.js 和特定于应用的 .js 文件）和加载项清单验证文件更新到版本 1.1。
ms.date: 12/04/2017
ms.openlocfilehash: c597c7456da2749d1061ab3e2c5bf9f41800a9cf
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944396"
---
# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a>更新到适用于 Office 的 JavaScript API 最新库和第 1.1 版加载项清单架构

本文介绍了如何将 Office 外接程序项目中的 JavaScript 文件（Office.js 和特定于应用程序的 .js 文件）和外接程序清单验证文件更新到版本 1.1。

## <a name="use-the-most-up-to-date-project-files"></a>使用最新项目文件

如果您使用 Visual Studio 来开发您的外接程序，以使用适用于 Office 的 JavaScript API 的 [最新 API 成员](https://docs.microsoft.com/javascript/office/what's-changed-in-the-javascript-api-for-office?view=office-js)和 [外接程序清单 v1.1 功能](../develop/add-in-manifests.md)（根据 offappmanifest-1.1.xsd 进行了验证），则您需要下载并安装 [Visual Studio 2015 和最新的 Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs)。

如果您使用文本编辑器或 Visual Studio 以外的 IDE 开发您的 外接程序，则您需要针对在 外接程序 的清单中引用的 Office.js 和架构版本，将引用更新到 CDN。

若要运行使用新的和已更新的 Office.js API 和加载项清单功能开发的加载项，您的客户必须运行 Office 2013 SP1 或更高版本的本地产品，并在适用的情况下运行 SharePoint Server 2013 SP1 和相关的服务器产品、Exchange Server 2013 Service Pack 1 (SP1) 或相当于联机托管的产品：Office 365、SharePoint Online 和 Exchange Online。

若要下载 Office、SharePoint 和 Exchange SP1 产品，请参阅以下内容：

- [Microsoft Office 2013 和相关桌面产品的所有 Service Pack 1 (SP1) 更新的列表](http://support.microsoft.com/kb/2850036)
    
- [Microsoft SharePoint Server 2013 和相关服务器产品的所有 Service Pack 1 (SP1) 更新的列表](http://support.microsoft.com/kb/2850035)
    
- [Exchange Server 2013 Service Pack 1 的说明](http://support.microsoft.com/kb/2926248)
    

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>更新使用 Visual Studio 创建的 Office 加载项项目

对于在适用于 Office 的 JavaScript API v1.1 和外接程序清单架构发布之前创建的项目，你可以使用“**NuGet 程序包管理器**”更新项目文件，然后更新外接程序的 HTML 页以进行引用。 

请注意，更新过程对于 _每个项目_ 执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，您需要重复更新过程。


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a>将项目中适用于 Office 的 JavaScript API 库文件更新到最新版本


1. 在 Visual Studio 2015 中，打开或新建“Office 加载项”**** 项目。
    
      - 在左侧窗格中，选择“**更新**”并完成程序包更新过程。
    
      - 转到步骤 6。
    
2. 依次选择“**工具**” > “**NuGet 包管理器**” > “**管理解决方案的 Nuget 包**”。
    
3. 在“**NuGet 程序包管理器**”中，为“**程序包源**”选择“**nuget.org**”并为“**筛选器**”选择“**可用升级**”。并选择 Microsoft.Office.js。
    
4. 在左侧窗格中，选择“更新”****，并完成包更新过程。
    
5. 在加载项 HTML 页面的 **head** 标记中，注释掉或删除任何现有 office.js 脚本引用，再引用更新后的适用于 Office 的 JavaScript API 库，如下所示：
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > 在 CDN URL 中，`/1/`  前面的`office.js`指定在第 1.1 版 Office.js 中使用最新增量版本。   


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>将项目中的清单文件更新为使用第 1.1 版架构

在外接程序清单文件中，更新 **OfficeApp**元素的 **xmlns**属性，将版本值更改为 `1.1`（除 **xmlns**属性以外的属性保持不变）。
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> 将加载项清单架构更新为第 1.1 版后，需要删除 **Capabilities**和 **Capability**元素，并将它们替换为 [Hosts](https://docs.microsoft.com/javascript/office/manifest/hosts?view=office-js)和 [Host](https://docs.microsoft.com/javascript/office/manifest/host?view=office-js)元素或 [Requirements 和 Requirement 元素](specify-office-hosts-and-api-requirements.md)。

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>更新使用文本编辑器或其他 IDE 创建的 Office 加载项项目

对于在发布适用于 Office 的 JavaScript API v1.1 和加载项清单架构之前创建的项目，您需要将加载项的 HTML 页更新到 v1.1 的 CDN 引用库中，将您的加载项清单文件更新为使用架构 v1.1。 

更新过程对_每个项目_分别执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，你需要重复更新过程。

你不需要适用于 Office 的 JavaScript API 文件（Office.js 和特定于应用程序的.js 文件）的本地副本来开发 Office 加载项（在运行时引用 Office.js 的 CDN 会下载必要的文件），但如果你想要库文件的本地副本，你可以使用 [NuGet 命令行实用程序](http://docs.nuget.org/consume/installing-nuget)和 `Install-Package Microsoft.Office.js` 命令来下载它们。

> [!NOTE] 
> 若要获取 v1.1 加载项清单的 XSD（XML 架构定义）副本，请参阅 [Office 加载项清单的架构参考（v1.1）](../develop/add-in-manifests.md)中列出的内容。


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a>将项目中适用于 Office 的 JavaScript API 库文件更新为使用最新版本

1. 在文本编辑器或 IDE 中，打开加载项 HTML 页面。
    
2. 在加载项 HTML 页面的 **head** 标记中，注释掉或删除任何现有 office.js 脚本引用，再引用更新后的适用于 Office 的 JavaScript API 库，如下所示：
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > 在 CDN URL 中，`/1/`前面的`office.js`指定在第 1.1 版 Office.js 中使用最新增量版本。   

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>将项目中的清单文件更新为使用第 1.1 版架构

在外接程序清单文件中，更新 **OfficeApp**元素的 **xmlns**属性，将版本值更改为 `1.1`（除 **xmlns**属性以外的属性保持不变）。
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> 将加载项清单架构更新为第 1.1 版后，需要删除 **Capabilities**和 **Capability**元素，并将它们替换为 [Hosts](https://docs.microsoft.com/javascript/office/manifest/hosts?view=office-js)和 [Host](https://docs.microsoft.com/javascript/office/manifest/host?view=office-js)元素或 [Requirements 和 Requirement 元素](specify-office-hosts-and-api-requirements.md)。
    

## <a name="see-also"></a>另请参阅

- [指定 Office 主机和 API 要求](specify-office-hosts-and-api-requirements.md) 
- [了解适用于 Office 的 JavaScript API](understanding-the-javascript-api-for-office.md)    
- [适用于 Office 的 JavaScript API](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)   
- [Office 外接程序清单的架构参考 (v1.1)](../develop/add-in-manifests.md)
    

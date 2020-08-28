---
title: 更新到最新的 Office JavaScript API 库和版本1.1 加载项清单架构
description: 将在 Office 加载项项目中的 JavaScript 文件（Office.js 和特定于应用程序的 .js 文件）和加载项清单验证文件更新到版本 1.1。
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: b0536b4b55accd99e002e26c467572330ba72ae2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293126"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>更新到最新的 Office JavaScript API 库和版本1.1 加载项清单架构

本文介绍了如何将 Office 外接程序项目中的 JavaScript 文件（Office.js 和特定于应用程序的 .js 文件）和外接程序清单验证文件更新到版本 1.1。

> [!NOTE]
> 在 Visual Studio 2019 中创建的项目已使用版本1.1。 但是，偶尔会对版本 1.1 进行次要更新，可使用本文中介绍的技术应用这些更新。

## <a name="use-the-most-up-to-date-project-files"></a>使用最新项目文件

如果使用 Visual Studio 开发外接程序，若要使用 Office JavaScript API 的最新 API 成员和 [外接程序清单的 v1.1 功能](../develop/add-in-manifests.md) (根据 offappmanifest-1.1) 进行验证，需要下载 Visual Studio 2019。 若要下载 Visual Studio 2019，请参阅 [Visual STUDIO IDE 页面](https://visualstudio.microsoft.com/vs/)。 在安装过程中，你需要选择 Office/SharePoint 开发工作负载。

如果您使用文本编辑器或 Visual Studio 以外的 IDE 开发您的 外接程序，则您需要针对在 外接程序 的清单中引用的 Office.js 和架构版本，将引用更新到 CDN。

若要运行使用新的和更新的 Office.js API 和外接程序清单功能开发的外接程序，客户必须运行 Office 2013 SP1 或更高版本的本地产品，以及 SharePoint server 2013 SP1 和相关服务器产品、Exchange Server 2013 Service Pack 1 (SP1) 或等效的 online 托管产品： Microsoft 365、SharePoint Online 和 Exchange Online。

若要下载 Office、SharePoint 和 Exchange SP1 产品，请参阅以下内容：

- [Microsoft Office 2013 和相关桌面产品的所有 Service Pack 1 (SP1) 更新的列表](https://support.microsoft.com/kb/2850036)

- [Microsoft SharePoint Server 2013 和相关服务器产品的所有 Service Pack 1 (SP1) 更新的列表](https://support.microsoft.com/kb/2850035)

- [Exchange Server 2013 Service Pack 1 的说明](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>更新使用 Visual Studio 创建的 Office 加载项项目

对于 Office JavaScript API 和外接程序清单架构的版本1.1 之前创建的项目，您可以使用 **NuGet 包管理器**更新项目的文件，然后更新外接程序的 HTML 页面以引用这些页面。 

请注意，更新过程对于 _每个项目_ 执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，您需要重复更新过程。

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>将项目中的 Office JavaScript API 库文件更新到最新版本
以下步骤将 Office.js 库文件更新到最新版本。 这些步骤使用 Visual Studio 2019，但它们与 Visual Studio 的早期版本类似。

1. 在 Visual Studio 2019 中，打开或创建新的 **Office 加载项** 项目。
2. 选择**工具**  >  **nuget 包管理器**  >  **管理用于解决方案的 NuGet 包**。
3. 选择“更新”**** 选项卡。
4. 选择 Microsoft.Office.js。 确保程序包源来自 **nuget.org**。
5. 在左窗格中，选择 " **安装** " 并完成程序包更新过程。

需要执行其他步骤才能完成更新。 在外接程序的 HTML 页面的 **头** 标记中，注释掉或删除任何现有的 office.js 脚本引用，并引用更新的 OFFICE JavaScript API 库，如下所示：

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > 在 CDN URL 中，`office.js` 中的 `/1/` 指定在第 1 版 Office.js 中使用最新增量版本。


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>将项目中的清单文件更新为使用第 1.1 版架构

在加载项清单文件中，更新 **OfficeApp** 元素的 **xmlns** 属性，将版本值更改为 `1.1`（除 **xmlns** 属性以外的属性保持不变）。

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> 将加载项清单架构的版本更新为1.1 之后，需要删除这些 **功能** 和 **功能** 元素，并将其替换为 [Hosts](../reference/manifest/hosts.md) 和 [Host](../reference/manifest/host.md) 元素或 [要求和要求元素](specify-office-hosts-and-api-requirements.md)。

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>更新使用文本编辑器或其他 IDE 创建的 Office 加载项项目

对于 Office JavaScript API 和外接程序清单架构的版本1.1 之前创建的项目，您需要更新加载项的 HTML 页面以引用 v1.1 库的 CDN，并将外接程序清单文件更新为使用架构 v1.1。 

更新过程对_每个项目_分别执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，你需要重复更新过程。

不需要 Office JavaScript API 文件的本地副本 ( # A0 和应用程序特定的 js 文件) 开发 Office 外接程序 (引用 CDN Office.js 下载运行时) 中所需的文件，但如果您需要库文件的本地副本，则可以使用 [NuGet 命令行实用程序](https://docs.nuget.org/consume/installing-nuget) 和 `Install-Package Microsoft.Office.js` 命令下载这些文件。

> [!NOTE]
> 若要获取有关 v1.1 加载项清单的 XSD（XML 架构定义）副本，请参阅 [Office 加载项清单的架构参考 (v1.1)](../develop/add-in-manifests.md) 中列出的内容。


### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>将项目中的 Office JavaScript API 库文件更新为使用最新版本

1. 在您的文本编辑器或 IDE 中打开您的加载项的 HTML 页。

2. 在外接程序的 HTML 页面的 **头** 标记中，注释掉或删除任何现有的 office.js 脚本引用，并引用更新的 OFFICE JavaScript API 库，如下所示：

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > 在 CDN URL 中，`office.js` 前面的 `/1/` 指定在第 1 版 Office.js 中使用最新增量版本。

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>将项目中的清单文件更新为使用第 1.1 版架构

在加载项清单文件中，更新 **OfficeApp** 元素的 **xmlns** 属性，将版本值更改为 `1.1`（除 **xmlns** 属性以外的属性保持不变）。

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> 将加载项清单架构的版本更新为1.1 之后，需要删除这些 **功能** 和 **功能** 元素，并将其替换为 [Hosts](../reference/manifest/hosts.md) 和 [Host](../reference/manifest/host.md) 元素或 [要求和要求元素](specify-office-hosts-and-api-requirements.md)。

## <a name="see-also"></a>另请参阅

- [指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md) ]
- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
- [Office 外接程序清单的架构参考 (v1.1)](../develop/add-in-manifests.md)

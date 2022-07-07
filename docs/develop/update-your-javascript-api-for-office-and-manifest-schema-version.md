---
title: 更新到最新的 Office JavaScript API 库和版本 1.1 加载项清单架构
description: 将在 Office 加载项项目中的 JavaScript 文件（Office.js 和特定于应用程序的 .js 文件）和加载项清单验证文件更新到版本 1.1。
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 32fcadb6a36ca540a799f8d6a5dfa671ee5e5de8
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660198"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>更新到最新的 Office JavaScript API 库和版本 1.1 加载项清单架构

本文介绍了如何将 Office 外接程序项目中的 JavaScript 文件（Office.js 和特定于应用程序的 .js 文件）和外接程序清单验证文件更新到版本 1.1。

> [!NOTE]
> 在 Visual Studio 2019 中创建的项目已使用版本 1.1。 但是，偶尔会对版本 1.1 进行次要更新，可使用本文中介绍的技术应用这些更新。

## <a name="use-the-most-up-to-date-project-files"></a>使用最新项目文件

如果使用 Visual Studio 开发外接程序，若要使用 Office JavaScript API 的最新 API 成员和 [加载项清单的 v1.1 功能](../develop/add-in-manifests.md) ， (针对 offappmanifest-1.1.xsd) 进行验证，则需要下载 Visual Studio 2019。 若要下载 Visual Studio 2019，请参阅 [Visual Studio IDE 页面](https://visualstudio.microsoft.com/vs/)。 在安装过程中，你需要选择 Office/SharePoint 开发工作负载。

如果使用除 Visual Studio 以外的文本编辑器或 IDE 来开发外接程序，则需要更新对内容分发网络的引用 (CDN) Office.js以及外接程序清单中引用的架构版本。

若要运行使用新的和更新的Office.js API 和外接程序清单功能开发的外接程序，客户必须运行 Office 2013 SP1 或更高版本的本地产品，如果适用，则 SharePoint Server 2013 SP1 和相关服务器产品、Exchange Server 2013 Service Pack 1 (SP1) 或等效的联机托管产品：Microsoft 365、SharePoint Online、 和Exchange Online。

若要下载 Office、SharePoint 和 Exchange SP1 产品，请参阅以下内容：

- [Microsoft Office 2013 和相关桌面产品的所有 Service Pack 1 (SP1) 更新的列表](https://support.microsoft.com/kb/2850036)

- [Microsoft SharePoint Server 2013 和相关服务器产品的所有 Service Pack 1 (SP1) 更新的列表](https://support.microsoft.com/kb/2850035)

- [Exchange Server 2013 Service Pack 1 的说明](https://support.microsoft.com/kb/2926248)

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>更新使用 Visual Studio 创建的 Office 加载项项目

对于在发布 Office JavaScript API 和外接程序清单架构的 v1.1 之前创建的项目，可以使用 **NuGet 包** 管理器更新项目的文件，然后更新外接程序的 HTML 页面以引用它们。

请注意，更新过程对于 _每个项目_ 执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，您需要重复更新过程。

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>将项目中的 Office JavaScript API 库文件更新为最新版本

以下步骤将Office.js库文件更新为最新版本。 这些步骤使用 Visual Studio 2019，但与以前版本的 Visual Studio 类似。

1. 在 Visual Studio 2019 中，打开或创建新的 **Office 外接程序** 项目。
2. 选择 **“工具** > **NuGet 包管理** > 器 **管理解决方案的 Nuget 包**”。
3. 选择“更新”选项卡。
4. 选择 Microsoft.Office.js。 确保包源来自 **nuget.org**。
5. 在左窗格中，选择 **“安装** ”并完成包更新过程。

需要执行其他步骤才能完成更新。 在外接程序的 HTML 页面的 **头** 标记中，注释掉或删除任何现有office.js脚本引用，并按如下所示引用更新的 Office JavaScript API 库：

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE]
   > 在 CDN URL 中，`office.js` 中的 `/1/` 指定在第 1 版 Office.js 中使用最新增量版本。

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>将项目中的清单文件更新为使用第 1.1 版架构

在外接程序的清单文件中，更新元素的 **\<OfficeApp\>** **xmlns** 属性，将版本值更改为`1.1` (使 **xmlns** 属性以外的属性保持不变) 。

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> 将加载项清单架构的版本更新为 1.1 后，需要删除 **功能** 和 **功能** 元素，并将其替换为 [主机](/javascript/api/manifest/hosts) 和 [主机](/javascript/api/manifest/host) 元素或 [“要求和要求”元素](specify-office-hosts-and-api-requirements.md)。

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>更新使用文本编辑器或其他 IDE 创建的 Office 加载项项目

对于在发布 Office JavaScript API 和外接程序清单架构的 v1.1 之前创建的项目，需要更新外接程序的 HTML 页面以引用 v1.1 库的 CDN，并更新外接程序的清单文件以使用架构 v1.1。

更新过程对 _每个项目_ 分别执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，你需要重复更新过程。

不需要 Office JavaScript API 文件 (Office.js 和特定于应用的.js文件的本地副本) 开发Office 外接程序 (引用 CDN 以Office.js在运行时) 下载所需的文件，但如果需要库文件的本地副本，则可以使用 [NuGet Command-Line实用工具](https://docs.nuget.org/consume/installing-nuget) 和 `Install-Package Microsoft.Office.js` 命令下载它们。

> [!NOTE]
> 若要获取有关 v1.1 加载项清单的 XSD（XML 架构定义）副本，请参阅 [Office 加载项清单的架构参考 (v1.1)](../develop/add-in-manifests.md) 中列出的内容。

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>更新项目中的 Office JavaScript API 库文件以使用最新版本

1. 在您的文本编辑器或 IDE 中打开您的加载项的 HTML 页。

2. 在外接程序的 HTML 页面的 **头** 标记中，注释掉或删除任何现有office.js脚本引用，并按如下所示引用更新的 Office JavaScript API 库：

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > 在 CDN URL 中，`office.js` 前面的 `/1/` 指定在第 1 版 Office.js 中使用最新增量版本。

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>将项目中的清单文件更新为使用第 1.1 版架构

在外接程序的清单文件中，更新元素的 **\<OfficeApp\>** **xmlns** 属性，将版本值更改为`1.1` (使 **xmlns** 属性以外的属性保持不变) 。

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> 将加载项清单架构的版本更新为 1.1 后，需要删除 **功能** 和 **功能** 元素，并将其替换为 [主机](/javascript/api/manifest/hosts) 和 [主机](/javascript/api/manifest/host) 元素或 [“要求和要求”元素](specify-office-hosts-and-api-requirements.md)。

## <a name="see-also"></a>另请参阅

- [指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md) ]
- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
- [Office 外接程序清单的架构参考 (v1.1)](../develop/add-in-manifests.md)

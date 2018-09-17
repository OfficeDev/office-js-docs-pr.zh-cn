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
# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a><span data-ttu-id="9e33d-103">更新到适用于 Office 的 JavaScript API 最新库和第 1.1 版加载项清单架构</span><span class="sxs-lookup"><span data-stu-id="9e33d-103">Update to the latest JavaScript API for Office library and version 1.1 add-in manifest schema</span></span>

<span data-ttu-id="9e33d-104">本文介绍了如何将 Office 外接程序项目中的 JavaScript 文件（Office.js 和特定于应用程序的 .js 文件）和外接程序清单验证文件更新到版本 1.1。</span><span class="sxs-lookup"><span data-stu-id="9e33d-104">This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.</span></span>

## <a name="use-the-most-up-to-date-project-files"></a><span data-ttu-id="9e33d-105">使用最新项目文件</span><span class="sxs-lookup"><span data-stu-id="9e33d-105">Use the most up-to-date project files</span></span>

<span data-ttu-id="9e33d-106">如果您使用 Visual Studio 来开发您的外接程序，以使用适用于 Office 的 JavaScript API 的 [最新 API 成员](https://docs.microsoft.com/javascript/office/what's-changed-in-the-javascript-api-for-office?view=office-js)和 [外接程序清单 v1.1 功能](../develop/add-in-manifests.md)（根据 offappmanifest-1.1.xsd 进行了验证），则您需要下载并安装 [Visual Studio 2015 和最新的 Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs)。</span><span class="sxs-lookup"><span data-stu-id="9e33d-106">If you use Visual Studio to develop your add-in, to use the [newest API members](https://docs.microsoft.com/javascript/office/what's-changed-in-the-javascript-api-for-office?view=office-js) of the JavaScript API for Office and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download and install the [Visual Studio 2015 and the latest Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs).</span></span>

<span data-ttu-id="9e33d-107">如果您使用文本编辑器或 Visual Studio 以外的 IDE 开发您的 外接程序，则您需要针对在 外接程序 的清单中引用的 Office.js 和架构版本，将引用更新到 CDN。</span><span class="sxs-lookup"><span data-stu-id="9e33d-107">If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the CDN for Office.js and the version of schema referenced in your add-in's manifest.</span></span>

<span data-ttu-id="9e33d-108">若要运行使用新的和已更新的 Office.js API 和加载项清单功能开发的加载项，您的客户必须运行 Office 2013 SP1 或更高版本的本地产品，并在适用的情况下运行 SharePoint Server 2013 SP1 和相关的服务器产品、Exchange Server 2013 Service Pack 1 (SP1) 或相当于联机托管的产品：Office 365、SharePoint Online 和 Exchange Online。</span><span class="sxs-lookup"><span data-stu-id="9e33d-108">To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Office 365, SharePoint Online, and Exchange Online.</span></span>

<span data-ttu-id="9e33d-109">若要下载 Office、SharePoint 和 Exchange SP1 产品，请参阅以下内容：</span><span class="sxs-lookup"><span data-stu-id="9e33d-109">To download Office, SharePoint, and Exchange SP1 products, see the following:</span></span>

- [<span data-ttu-id="9e33d-110">Microsoft Office 2013 和相关桌面产品的所有 Service Pack 1 (SP1) 更新的列表</span><span class="sxs-lookup"><span data-stu-id="9e33d-110">List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products</span></span>](http://support.microsoft.com/kb/2850036)
    
- [<span data-ttu-id="9e33d-111">Microsoft SharePoint Server 2013 和相关服务器产品的所有 Service Pack 1 (SP1) 更新的列表</span><span class="sxs-lookup"><span data-stu-id="9e33d-111">List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products</span></span>](http://support.microsoft.com/kb/2850035)
    
- [<span data-ttu-id="9e33d-112">Exchange Server 2013 Service Pack 1 的说明</span><span class="sxs-lookup"><span data-stu-id="9e33d-112">Description of Exchange Server 2013 Service Pack 1</span></span>](http://support.microsoft.com/kb/2926248)
    

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a><span data-ttu-id="9e33d-113">更新使用 Visual Studio 创建的 Office 加载项项目</span><span class="sxs-lookup"><span data-stu-id="9e33d-113">Updating an Office Add-in project created with Visual Studio</span></span>

<span data-ttu-id="9e33d-114">对于在适用于 Office 的 JavaScript API v1.1 和外接程序清单架构发布之前创建的项目，你可以使用“**NuGet 程序包管理器**”更新项目文件，然后更新外接程序的 HTML 页以进行引用。</span><span class="sxs-lookup"><span data-stu-id="9e33d-114">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you can update a project's files using the  **NuGet Package Manager**, and then update your add-in's HTML pages to reference them.</span></span> 

<span data-ttu-id="9e33d-115">请注意，更新过程对于 _每个项目_ 执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，您需要重复更新过程。</span><span class="sxs-lookup"><span data-stu-id="9e33d-115">Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a><span data-ttu-id="9e33d-116">将项目中适用于 Office 的 JavaScript API 库文件更新到最新版本</span><span class="sxs-lookup"><span data-stu-id="9e33d-116">Update the JavaScript API for Office library files in your project to the newest release</span></span>


1. <span data-ttu-id="9e33d-117">在 Visual Studio 2015 中，打开或新建“Office 加载项”\*\*\*\* 项目。</span><span class="sxs-lookup"><span data-stu-id="9e33d-117">In Visual Studio 2015, open or create a new  **Office Add-in** project.</span></span>
    
      - <span data-ttu-id="9e33d-118">在左侧窗格中，选择“**更新**”并完成程序包更新过程。</span><span class="sxs-lookup"><span data-stu-id="9e33d-118">In the left pane, choose **Update** and complete the package update process.</span></span>
    
      - <span data-ttu-id="9e33d-119">转到步骤 6。</span><span class="sxs-lookup"><span data-stu-id="9e33d-119">Go to step 6.</span></span>
    
2. <span data-ttu-id="9e33d-120">依次选择“**工具**” > “**NuGet 包管理器**” > “**管理解决方案的 Nuget 包**”。</span><span class="sxs-lookup"><span data-stu-id="9e33d-120">Choose  **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.</span></span>
    
3. <span data-ttu-id="9e33d-p101">在“**NuGet 程序包管理器**”中，为“**程序包源**”选择“**nuget.org**”并为“**筛选器**”选择“**可用升级**”。并选择 Microsoft.Office.js。</span><span class="sxs-lookup"><span data-stu-id="9e33d-p101">In the  **NuGet Package Manager**, select  **nuget.org** for **Package source** and **Upgrade available** for **Filter**. and select Microsoft.Office.js.</span></span>
    
4. <span data-ttu-id="9e33d-123">在左侧窗格中，选择“更新”\*\*\*\*，并完成包更新过程。</span><span class="sxs-lookup"><span data-stu-id="9e33d-123">In the left pane, choose **Update** and complete the package update process.</span></span>
    
5. <span data-ttu-id="9e33d-124">在加载项 HTML 页面的 **head** 标记中，注释掉或删除任何现有 office.js 脚本引用，再引用更新后的适用于 Office 的 JavaScript API 库，如下所示：</span><span class="sxs-lookup"><span data-stu-id="9e33d-124">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > <span data-ttu-id="9e33d-125">在 CDN URL 中，`/1/`  前面的`office.js`指定在第 1.1 版 Office.js 中使用最新增量版本。</span><span class="sxs-lookup"><span data-stu-id="9e33d-125">NOTE The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>   


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="9e33d-126">将项目中的清单文件更新为使用第 1.1 版架构</span><span class="sxs-lookup"><span data-stu-id="9e33d-126">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="9e33d-127">在外接程序清单文件中，更新 **OfficeApp**元素的 **xmlns**属性，将版本值更改为 `1.1`（除 **xmlns**属性以外的属性保持不变）。</span><span class="sxs-lookup"><span data-stu-id="9e33d-127">In your Add-in's Manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> <span data-ttu-id="9e33d-128">将加载项清单架构更新为第 1.1 版后，需要删除 **Capabilities**和 **Capability**元素，并将它们替换为 [Hosts](https://docs.microsoft.com/javascript/office/manifest/hosts?view=office-js)和 [Host](https://docs.microsoft.com/javascript/office/manifest/host?view=office-js)元素或 [Requirements 和 Requirement 元素](specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="9e33d-128">After updating the version of the add-in manifest schema to 1.1, you will need to remove the Capabilities and Capability elements, and replace them with either the Hosts and Host elements or the Requirements and Requirement elements.</span></span>

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a><span data-ttu-id="9e33d-129">更新使用文本编辑器或其他 IDE 创建的 Office 加载项项目</span><span class="sxs-lookup"><span data-stu-id="9e33d-129">Updating an Office Add-in project created with a text editor or other IDE</span></span>

<span data-ttu-id="9e33d-130">对于在发布适用于 Office 的 JavaScript API v1.1 和加载项清单架构之前创建的项目，您需要将加载项的 HTML 页更新到 v1.1 的 CDN 引用库中，将您的加载项清单文件更新为使用架构 v1.1。</span><span class="sxs-lookup"><span data-stu-id="9e33d-130">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1.</span></span> 

<span data-ttu-id="9e33d-131">更新过程对_每个项目_分别执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，你需要重复更新过程。</span><span class="sxs-lookup"><span data-stu-id="9e33d-131">The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

<span data-ttu-id="9e33d-132">你不需要适用于 Office 的 JavaScript API 文件（Office.js 和特定于应用程序的.js 文件）的本地副本来开发 Office 加载项（在运行时引用 Office.js 的 CDN 会下载必要的文件），但如果你想要库文件的本地副本，你可以使用 [NuGet 命令行实用程序](http://docs.nuget.org/consume/installing-nuget)和 `Install-Package Microsoft.Office.js` 命令来下载它们。</span><span class="sxs-lookup"><span data-stu-id="9e33d-132">You don't need local copies of the JavaScript API for Office files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](http://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.</span></span>

> [!NOTE] 
> <span data-ttu-id="9e33d-133">若要获取 v1.1 加载项清单的 XSD（XML 架构定义）副本，请参阅 [Office 加载项清单的架构参考（v1.1）](../develop/add-in-manifests.md)中列出的内容。</span><span class="sxs-lookup"><span data-stu-id="9e33d-133">NOTE To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a><span data-ttu-id="9e33d-134">将项目中适用于 Office 的 JavaScript API 库文件更新为使用最新版本</span><span class="sxs-lookup"><span data-stu-id="9e33d-134">Update the JavaScript API for Office library files in your project to use the newest release</span></span>

1. <span data-ttu-id="9e33d-135">在文本编辑器或 IDE 中，打开加载项 HTML 页面。</span><span class="sxs-lookup"><span data-stu-id="9e33d-135">Open the HTML pages for your add-in in your text editor or IDE.</span></span>
    
2. <span data-ttu-id="9e33d-136">在加载项 HTML 页面的 **head** 标记中，注释掉或删除任何现有 office.js 脚本引用，再引用更新后的适用于 Office 的 JavaScript API 库，如下所示：</span><span class="sxs-lookup"><span data-stu-id="9e33d-136">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > <span data-ttu-id="9e33d-137">在 CDN URL 中，`/1/`前面的`office.js`指定在第 1.1 版 Office.js 中使用最新增量版本。</span><span class="sxs-lookup"><span data-stu-id="9e33d-137">NOTE The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>   

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="9e33d-138">将项目中的清单文件更新为使用第 1.1 版架构</span><span class="sxs-lookup"><span data-stu-id="9e33d-138">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="9e33d-139">在外接程序清单文件中，更新 **OfficeApp**元素的 **xmlns**属性，将版本值更改为 `1.1`（除 **xmlns**属性以外的属性保持不变）。</span><span class="sxs-lookup"><span data-stu-id="9e33d-139">In your Add-in's Manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> <span data-ttu-id="9e33d-140">将加载项清单架构更新为第 1.1 版后，需要删除 **Capabilities**和 **Capability**元素，并将它们替换为 [Hosts](https://docs.microsoft.com/javascript/office/manifest/hosts?view=office-js)和 [Host](https://docs.microsoft.com/javascript/office/manifest/host?view=office-js)元素或 [Requirements 和 Requirement 元素](specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="9e33d-140">After updating the version of the add-in manifest schema to 1.1, you will need to remove the Capabilities and Capability elements, and replace them with either the Hosts and Host elements or the Requirements and Requirement elements.</span></span>
    

## <a name="see-also"></a><span data-ttu-id="9e33d-141">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9e33d-141">See also</span></span>

- [<span data-ttu-id="9e33d-142">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="9e33d-142">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md) 
- [<span data-ttu-id="9e33d-143">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="9e33d-143">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)    
- [<span data-ttu-id="9e33d-144">适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="9e33d-144">JavaScript API for Office</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)   
- [<span data-ttu-id="9e33d-145">Office 外接程序清单的架构参考 (v1.1)</span><span class="sxs-lookup"><span data-stu-id="9e33d-145">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
    

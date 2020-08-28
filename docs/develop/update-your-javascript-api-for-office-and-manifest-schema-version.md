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
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a><span data-ttu-id="d5851-103">更新到最新的 Office JavaScript API 库和版本1.1 加载项清单架构</span><span class="sxs-lookup"><span data-stu-id="d5851-103">Update to the latest Office JavaScript API library and version 1.1 add-in manifest schema</span></span>

<span data-ttu-id="d5851-104">本文介绍了如何将 Office 外接程序项目中的 JavaScript 文件（Office.js 和特定于应用程序的 .js 文件）和外接程序清单验证文件更新到版本 1.1。</span><span class="sxs-lookup"><span data-stu-id="d5851-104">This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="d5851-105">在 Visual Studio 2019 中创建的项目已使用版本1.1。</span><span class="sxs-lookup"><span data-stu-id="d5851-105">Projects created in Visual Studio 2019 will already use version 1.1.</span></span> <span data-ttu-id="d5851-106">但是，偶尔会对版本 1.1 进行次要更新，可使用本文中介绍的技术应用这些更新。</span><span class="sxs-lookup"><span data-stu-id="d5851-106">However there are occasional minor updates to version 1.1 that you can apply by using the techniques in this article.</span></span>

## <a name="use-the-most-up-to-date-project-files"></a><span data-ttu-id="d5851-107">使用最新项目文件</span><span class="sxs-lookup"><span data-stu-id="d5851-107">Use the most up-to-date project files</span></span>

<span data-ttu-id="d5851-108">如果使用 Visual Studio 开发外接程序，若要使用 Office JavaScript API 的最新 API 成员和 [外接程序清单的 v1.1 功能](../develop/add-in-manifests.md) (根据 offappmanifest-1.1) 进行验证，需要下载 Visual Studio 2019。</span><span class="sxs-lookup"><span data-stu-id="d5851-108">If you use Visual Studio to develop your add-in, to use the newest API members of the Office JavaScript API and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download Visual Studio 2019.</span></span> <span data-ttu-id="d5851-109">若要下载 Visual Studio 2019，请参阅 [Visual STUDIO IDE 页面](https://visualstudio.microsoft.com/vs/)。</span><span class="sxs-lookup"><span data-stu-id="d5851-109">To download Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="d5851-110">在安装过程中，你需要选择 Office/SharePoint 开发工作负载。</span><span class="sxs-lookup"><span data-stu-id="d5851-110">During installation you'll need to select the Office/SharePoint development workload.</span></span>

<span data-ttu-id="d5851-111">如果您使用文本编辑器或 Visual Studio 以外的 IDE 开发您的 外接程序，则您需要针对在 外接程序 的清单中引用的 Office.js 和架构版本，将引用更新到 CDN。</span><span class="sxs-lookup"><span data-stu-id="d5851-111">If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the CDN for Office.js and the version of schema referenced in your add-in's manifest.</span></span>

<span data-ttu-id="d5851-112">若要运行使用新的和更新的 Office.js API 和外接程序清单功能开发的外接程序，客户必须运行 Office 2013 SP1 或更高版本的本地产品，以及 SharePoint server 2013 SP1 和相关服务器产品、Exchange Server 2013 Service Pack 1 (SP1) 或等效的 online 托管产品： Microsoft 365、SharePoint Online 和 Exchange Online。</span><span class="sxs-lookup"><span data-stu-id="d5851-112">To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Microsoft 365, SharePoint Online, and Exchange Online.</span></span>

<span data-ttu-id="d5851-113">若要下载 Office、SharePoint 和 Exchange SP1 产品，请参阅以下内容：</span><span class="sxs-lookup"><span data-stu-id="d5851-113">To download Office, SharePoint, and Exchange SP1 products, see the following:</span></span>

- [<span data-ttu-id="d5851-114">Microsoft Office 2013 和相关桌面产品的所有 Service Pack 1 (SP1) 更新的列表</span><span class="sxs-lookup"><span data-stu-id="d5851-114">List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products</span></span>](https://support.microsoft.com/kb/2850036)

- [<span data-ttu-id="d5851-115">Microsoft SharePoint Server 2013 和相关服务器产品的所有 Service Pack 1 (SP1) 更新的列表</span><span class="sxs-lookup"><span data-stu-id="d5851-115">List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products</span></span>](https://support.microsoft.com/kb/2850035)

- [<span data-ttu-id="d5851-116">Exchange Server 2013 Service Pack 1 的说明</span><span class="sxs-lookup"><span data-stu-id="d5851-116">Description of Exchange Server 2013 Service Pack 1</span></span>](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a><span data-ttu-id="d5851-117">更新使用 Visual Studio 创建的 Office 加载项项目</span><span class="sxs-lookup"><span data-stu-id="d5851-117">Updating an Office Add-in project created with Visual Studio</span></span>

<span data-ttu-id="d5851-118">对于 Office JavaScript API 和外接程序清单架构的版本1.1 之前创建的项目，您可以使用 **NuGet 包管理器**更新项目的文件，然后更新外接程序的 HTML 页面以引用这些页面。</span><span class="sxs-lookup"><span data-stu-id="d5851-118">For projects created before the release of v1.1 of the Office JavaScript API and add-in manifest schema, you can update a project's files using the **NuGet Package Manager**, and then update your add-in's HTML pages to reference them.</span></span> 

<span data-ttu-id="d5851-119">请注意，更新过程对于 _每个项目_ 执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，您需要重复更新过程。</span><span class="sxs-lookup"><span data-stu-id="d5851-119">Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a><span data-ttu-id="d5851-120">将项目中的 Office JavaScript API 库文件更新到最新版本</span><span class="sxs-lookup"><span data-stu-id="d5851-120">Update the Office JavaScript API library files in your project to the newest release</span></span>
<span data-ttu-id="d5851-121">以下步骤将 Office.js 库文件更新到最新版本。</span><span class="sxs-lookup"><span data-stu-id="d5851-121">The following steps will update your Office.js library files to the latest version.</span></span> <span data-ttu-id="d5851-122">这些步骤使用 Visual Studio 2019，但它们与 Visual Studio 的早期版本类似。</span><span class="sxs-lookup"><span data-stu-id="d5851-122">The steps use Visual Studio 2019, but they are similar for previous versions of Visual Studio.</span></span>

1. <span data-ttu-id="d5851-123">在 Visual Studio 2019 中，打开或创建新的 **Office 加载项** 项目。</span><span class="sxs-lookup"><span data-stu-id="d5851-123">In Visual Studio 2019, open or create a new **Office Add-in** project.</span></span>
2. <span data-ttu-id="d5851-124">选择**工具**  >  **nuget 包管理器**  >  **管理用于解决方案的 NuGet 包**。</span><span class="sxs-lookup"><span data-stu-id="d5851-124">Choose **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.</span></span>
3. <span data-ttu-id="d5851-125">选择“更新”\*\*\*\* 选项卡。</span><span class="sxs-lookup"><span data-stu-id="d5851-125">Choose the **Updates** tab.</span></span>
4. <span data-ttu-id="d5851-126">选择 Microsoft.Office.js。</span><span class="sxs-lookup"><span data-stu-id="d5851-126">Select Microsoft.Office.js.</span></span> <span data-ttu-id="d5851-127">确保程序包源来自 **nuget.org**。</span><span class="sxs-lookup"><span data-stu-id="d5851-127">Ensure the package source is from **nuget.org**.</span></span>
5. <span data-ttu-id="d5851-128">在左窗格中，选择 " **安装** " 并完成程序包更新过程。</span><span class="sxs-lookup"><span data-stu-id="d5851-128">In the left pane, choose **Install** and complete the package update process.</span></span>

<span data-ttu-id="d5851-129">需要执行其他步骤才能完成更新。</span><span class="sxs-lookup"><span data-stu-id="d5851-129">You'll need to take a few additional steps to complete the update.</span></span> <span data-ttu-id="d5851-130">在外接程序的 HTML 页面的 **头** 标记中，注释掉或删除任何现有的 office.js 脚本引用，并引用更新的 OFFICE JavaScript API 库，如下所示：</span><span class="sxs-lookup"><span data-stu-id="d5851-130">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated Office JavaScript API library as follows:</span></span>

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > <span data-ttu-id="d5851-131">在 CDN URL 中，`office.js` 中的 `/1/` 指定在第 1 版 Office.js 中使用最新增量版本。</span><span class="sxs-lookup"><span data-stu-id="d5851-131">The `/1/` in the `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="d5851-132">将项目中的清单文件更新为使用第 1.1 版架构</span><span class="sxs-lookup"><span data-stu-id="d5851-132">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="d5851-133">在加载项清单文件中，更新 **OfficeApp** 元素的 **xmlns** 属性，将版本值更改为 `1.1`（除 **xmlns** 属性以外的属性保持不变）。</span><span class="sxs-lookup"><span data-stu-id="d5851-133">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="d5851-134">将加载项清单架构的版本更新为1.1 之后，需要删除这些 **功能** 和 **功能** 元素，并将其替换为 [Hosts](../reference/manifest/hosts.md) 和 [Host](../reference/manifest/host.md) 元素或 [要求和要求元素](specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="d5851-134">After updating the version of the add-in manifest schema to 1.1, you will need to remove the **Capabilities** and **Capability** elements, and replace them with either the [Hosts](../reference/manifest/hosts.md) and [Host](../reference/manifest/host.md) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a><span data-ttu-id="d5851-135">更新使用文本编辑器或其他 IDE 创建的 Office 加载项项目</span><span class="sxs-lookup"><span data-stu-id="d5851-135">Updating an Office Add-in project created with a text editor or other IDE</span></span>

<span data-ttu-id="d5851-136">对于 Office JavaScript API 和外接程序清单架构的版本1.1 之前创建的项目，您需要更新加载项的 HTML 页面以引用 v1.1 库的 CDN，并将外接程序清单文件更新为使用架构 v1.1。</span><span class="sxs-lookup"><span data-stu-id="d5851-136">For projects created before the release of v1.1 of the Office JavaScript API and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1.</span></span> 

<span data-ttu-id="d5851-137">更新过程对_每个项目_分别执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，你需要重复更新过程。</span><span class="sxs-lookup"><span data-stu-id="d5851-137">The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

<span data-ttu-id="d5851-138">不需要 Office JavaScript API 文件的本地副本 ( # A0 和应用程序特定的 js 文件) 开发 Office 外接程序 (引用 CDN Office.js 下载运行时) 中所需的文件，但如果您需要库文件的本地副本，则可以使用 [NuGet 命令行实用程序](https://docs.nuget.org/consume/installing-nuget) 和 `Install-Package Microsoft.Office.js` 命令下载这些文件。</span><span class="sxs-lookup"><span data-stu-id="d5851-138">You don't need local copies of the Office JavaScript API files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](https://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.</span></span>

> [!NOTE]
> <span data-ttu-id="d5851-139">若要获取有关 v1.1 加载项清单的 XSD（XML 架构定义）副本，请参阅 [Office 加载项清单的架构参考 (v1.1)](../develop/add-in-manifests.md) 中列出的内容。</span><span class="sxs-lookup"><span data-stu-id="d5851-139">To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>


### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a><span data-ttu-id="d5851-140">将项目中的 Office JavaScript API 库文件更新为使用最新版本</span><span class="sxs-lookup"><span data-stu-id="d5851-140">Update the Office JavaScript API library files in your project to use the newest release</span></span>

1. <span data-ttu-id="d5851-141">在您的文本编辑器或 IDE 中打开您的加载项的 HTML 页。</span><span class="sxs-lookup"><span data-stu-id="d5851-141">Open the HTML pages for your add-in in your text editor or IDE.</span></span>

2. <span data-ttu-id="d5851-142">在外接程序的 HTML 页面的 **头** 标记中，注释掉或删除任何现有的 office.js 脚本引用，并引用更新的 OFFICE JavaScript API 库，如下所示：</span><span class="sxs-lookup"><span data-stu-id="d5851-142">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated Office JavaScript API library as follows:</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > <span data-ttu-id="d5851-143">在 CDN URL 中，`office.js` 前面的 `/1/` 指定在第 1 版 Office.js 中使用最新增量版本。</span><span class="sxs-lookup"><span data-stu-id="d5851-143">The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="d5851-144">将项目中的清单文件更新为使用第 1.1 版架构</span><span class="sxs-lookup"><span data-stu-id="d5851-144">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="d5851-145">在加载项清单文件中，更新 **OfficeApp** 元素的 **xmlns** 属性，将版本值更改为 `1.1`（除 **xmlns** 属性以外的属性保持不变）。</span><span class="sxs-lookup"><span data-stu-id="d5851-145">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="d5851-146">将加载项清单架构的版本更新为1.1 之后，需要删除这些 **功能** 和 **功能** 元素，并将其替换为 [Hosts](../reference/manifest/hosts.md) 和 [Host](../reference/manifest/host.md) 元素或 [要求和要求元素](specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="d5851-146">After updating the version of the add-in manifest schema to 1.1, you will need to remove the **Capabilities** and **Capability** elements, and replace them with either the [Hosts](../reference/manifest/hosts.md) and [Host](../reference/manifest/host.md) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d5851-147">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d5851-147">See also</span></span>

- <span data-ttu-id="d5851-148">[指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md) ]</span><span class="sxs-lookup"><span data-stu-id="d5851-148">[Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md) ]</span></span>
- [<span data-ttu-id="d5851-149">了解 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="d5851-149">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="d5851-150">Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="d5851-150">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="d5851-151">Office 外接程序清单的架构参考 (v1.1)</span><span class="sxs-lookup"><span data-stu-id="d5851-151">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)

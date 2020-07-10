---
title: 在 Microsoft Azure 上托管 Office 加载项 | Microsoft Docs
description: 了解如何将加载项 Web 应用部署到 Azure 并旁加载该加载项以便在 Office 客户端应用程序中进行测试。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: a30f1a8219501a68e6f46f013ef46640a59fe4e9
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094230"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a>在 Microsoft Azure 上托管 Office 加载项

The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office desktop applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.

本文介绍了如何将外接程序 Web 应用部署到 Azure 并[旁加载外接程序](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)以在 Office 客户端应用程序中进行测试。

## <a name="prerequisites"></a>先决条件 

1. 安装 [Visual Studio 2019](https://www.visualstudio.com/downloads)，并选择添加 **Azure 开发**工作负载。

    > [!NOTE]
    > 如果之前已安装 Visual Studio 2019，请[使用 Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)，以确保安装 **Azure 开发**工作负载。 

2. 安装 Office。

    > [!NOTE]
    > 如果尚未安装 Office，可以[注册 1 个月免费试用版](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)。

3. 获取 Azure 订阅。

    > [!NOTE]
    > 如果还没有 Azure 订阅，可以[通过 Visual Studio 订阅获取 Azure 订阅](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/)，也可以[注册免费试用版](https://azure.microsoft.com/pricing/free-trial)。 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a>第 1 步：创建用于托管加载项 XML 清单文件的共享文件夹

1. 打开开发计算机的文件资源管理器。

2. 右键单击 C:\ 驱动器，然后选择“新建”**** > “文件夹”****。

3. 将新文件夹命名为 AddinManifests。

4. 右键单击 AddinManifests 文件夹，然后选择“共享”**** > “特定用户”****。

5. 在“文件共享”**** 中，选择下拉箭头，再依次选择“所有人”**** > “添加”**** > “共享”****。

> [!NOTE]
> In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a>第 2 步：将文件共享添加到受信任的加载项目录

1. 启动 Word 并创建文档。

    > [!NOTE]
    > 尽管本示例使用的是 Word，但也可以使用任何支持 Office 加载项的 Office 应用（如 Excel、Outlook、PowerPoint 或 Project）。

2. 选择“**文件**” > “**选项**”。

3. 在“Word 选项”**** 对话框中，选择“信任中心”****，然后选择“信任中心设置”****。

4. In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**. 

5. 选中“在菜单中显示”**** 复选框。

    > [!NOTE]
    > 如果将加载项 XML 清单文件存储到已指定为受信任的 Web 加载项目录的共享中，用户可以转到功能区中的“插入”**** 选项卡，并选择“我的加载项”****，此时加载项就会显示在“Office 加载项”**** 对话框中的“共享文件夹”**** 下。

6. 关闭 Word。

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a>第 3 步：使用 Azure 门户在 Azure 中创建 Web 应用

若要使用 Azure 门户创建 Web 应用，请完成以下步骤。

1. 使用 Azure 凭据登录到 [Azure 门户](https://portal.azure.com/)。

2. 在“**Azure 服务**”下，选择“**Web 应用**”。

3. 在“**应用服务**”页面上，选择“**添加**”。 提供以下信息：

      - 选择要用于创建此站点的“订阅”****。
      
      - Choose the **Resource Group** for your site. If you create a new group, you also need to name it.
      
      - 为站点输入唯一的“应用名称”****。 Azure 验证站点名称在整个 azureweb apps.net 域中是否是唯一的。

      - 选择使用代码还是 Docker 容器进行发布。

      - 指定“**运行时堆栈**”。

      - 为站点选择“**操作系统**”。

      - 选择“**区域**”。

      - 选择要用于创建此站点的“**应用服务计划**”。

      - 选择“**创建**”。

4. 下一页将显示部署的进行状态和完成时间。 完成后，选择“**转到资源**”。  

5. 在“**概述**”节中，选择在“**URL**”下显示的 URL。 随即将打开浏览器，并显示包含“应用服务应用已启动且正在运行”消息的网页。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure 网站自动提供 HTTPS 终结点。

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a>第 4 步：在 Visual Studio 中创建 Office 加载项

1. 以管理员身份启动 Visual Studio。

2. 选择“**创建新项目**”。

3. 使用搜索框，输入“**加载程序**”。

4. 选择“**Word Web 外接程序**”作为项目类型，然后选择“**下一步**”以接受默认设置。

Visual Studio 将创建基本的 Word 外接程序，你可以按原样发布，无需对其 Web 项目进行任何更改。 若要为其他 Office 主机类型（如 Excel）生成外接程序，请重复这些步骤，并选择具有所需 Office 主机的项目类型。

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a>第 5 步：将 Office 外接程序 Web 应用发布到 Azure

1. 在 Visual Studio 中打开外接程序项目后，展开“**解决方案资源管理器**”中的解决方案节点，然后选择“**应用服务**”。

2. Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.

3. 在“发布”**** 选项卡上：

      - 选择“Microsoft Azure 应用服务”****。

      - 选择“选择现有”****。

      - 选择“发布”****。

4. Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.

5. 复制根 URL（例如：https://YourDomain.azurewebsites.net)；本文后续部分中编辑加载项清单文件时将需要此 URL。

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a>第 6 步：编辑并部署加载项 XML 清单文件

1. 在示例 Office 外接程序在“解决方案资源管理器”**** 中打开的 Visual Studio 中，展开该解决方案以显示两个项目。

2. Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.

3. In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net). 

4. 选择“**文件**”，然后选择“**全部保存**”。 然后复制外接程序 XML 清单文件（例如 WordWebAddIn.xml）。

5. 使用“**文件资源管理器**”程序浏览到在[第 1 步：创建共享文件夹](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file)中创建的网络文件共享，并将清单文件粘贴到此文件夹。

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a>第 7 步：在 Office 客户端应用程序中插入并运行加载项

1. 启动 Word 并创建文档。

2. 在功能区中选择“**插入**” > “**我的加载项**”。

3. In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.

4. Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.

5. On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.

6. Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.

## <a name="see-also"></a>另请参阅

- [发布 Office 加载项](../publish/publish.md)
- [使用 Visual Studio 发布加载项](../publish/package-your-add-in-using-visual-studio.md)

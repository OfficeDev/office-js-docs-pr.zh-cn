---
title: 安装最新版 Office 2016
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 98dc69a7971a94b96bc3f7304fc7905f31013a87
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925232"
---
# <a name="install-the-latest-version-of-office-2016"></a><span data-ttu-id="b2720-102">安装最新版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="b2720-102">Install the latest version of Office 2016</span></span>

<span data-ttu-id="b2720-103">新开发人员功能（包括仍处于预览阶段的功能）会先向选择获取最新版 Office 的订阅者提供。</span><span class="sxs-lookup"><span data-stu-id="b2720-103">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="b2720-104">选择获取最新版</span><span class="sxs-lookup"><span data-stu-id="b2720-104">Opt in to getting the latest builds</span></span>

<span data-ttu-id="b2720-105">若要选择获取最新版 Office 2016，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="b2720-105">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="b2720-106">如果是 Office 365 家庭版、个人版或大专院校版订阅者，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider)。</span><span class="sxs-lookup"><span data-stu-id="b2720-106">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="b2720-107">如果你是 Office 365 商业版客户，请参阅 [为 Office 365 商业版客户安装首次发布](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)。</span><span class="sxs-lookup"><span data-stu-id="b2720-107">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="b2720-108">如果在 Mac 上运行 Office 2016：</span><span class="sxs-lookup"><span data-stu-id="b2720-108">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="b2720-109">启动 Office 2016 for Mac 程序。</span><span class="sxs-lookup"><span data-stu-id="b2720-109">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="b2720-110">选择“帮助”菜单上的“**检查更新**”。</span><span class="sxs-lookup"><span data-stu-id="b2720-110">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="b2720-111">选中“Microsoft 自动更新”框，以加入 Office 预览体验成员计划。</span><span class="sxs-lookup"><span data-stu-id="b2720-111">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="b2720-112">获取最新版</span><span class="sxs-lookup"><span data-stu-id="b2720-112">Get the latest build</span></span>

<span data-ttu-id="b2720-113">若要获取最新版 Office 2016，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="b2720-113">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="b2720-114">下载 [Office 2016 部署工具](https://www.microsoft.com/download/details.aspx?id=49117)。</span><span class="sxs-lookup"><span data-stu-id="b2720-114">Download the [Office 2016 Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span> 
2. <span data-ttu-id="b2720-p101">运行该工具。这会提取以下两个文件：Setup.exe 和 configuration.xml。</span><span class="sxs-lookup"><span data-stu-id="b2720-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="b2720-117">将 configuration.xml 文件替换为[首次发布配置文件](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)。</span><span class="sxs-lookup"><span data-stu-id="b2720-117">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="b2720-118">以管理员身份运行下面的命令：`setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="b2720-118">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="b2720-119">此命令可能需要运行很长时间才能完成，而且不会显示进度。</span><span class="sxs-lookup"><span data-stu-id="b2720-119">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="b2720-p102">在安装进程完成后，你已安装最新的 Office 2016 应用程序。要验证你是否拥有最新版本，请从任何 Office 应用程序转到“**文件**” > “**帐户**”。在“Office 更新”下，你将看到版本号上面的 (Office Insiders) 标签。</span><span class="sxs-lookup"><span data-stu-id="b2720-p102">When the installation process finishes, you will have the latest Office 2016 applications installed. To verify that you have the latest build, go to **File** > **Account** from any Office application. Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![显示产品信息的屏幕截图（带有 Office Insiders 标签）](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="b2720-124">Office JavaScript API 要求集对应的最低 Office 内部版本</span><span class="sxs-lookup"><span data-stu-id="b2720-124">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="b2720-125">若要了解 API 要求集对应的各个平台的最低产品内部版本，请参阅以下资源：</span><span class="sxs-lookup"><span data-stu-id="b2720-125">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="b2720-126">Word JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="b2720-126">Word JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets)
- [<span data-ttu-id="b2720-127">Excel JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="b2720-127">Excel JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets)
- [<span data-ttu-id="b2720-128">OneNote JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="b2720-128">OneNote JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets)
- [<span data-ttu-id="b2720-129">对话框 API 要求集</span><span class="sxs-lookup"><span data-stu-id="b2720-129">Dialog API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="b2720-130">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="b2720-130">Office common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)

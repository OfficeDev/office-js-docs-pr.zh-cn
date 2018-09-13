---
title: 安装最新版本 Office
description: 有关如何选择获取 Office 最新版本的信息。
ms.date: 12/04/2017
ms.openlocfilehash: 14e26d9fa9f7ec3b2724cbf2e9787cde9dbe4094
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943878"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="42e15-103">安装最新版本 Office</span><span class="sxs-lookup"><span data-stu-id="42e15-103">Install the latest version of Office</span></span>

<span data-ttu-id="42e15-104">新开发人员功能（包括仍处于预览阶段的功能）会首先向选择获取最新版 Office 的订阅者提供。</span><span class="sxs-lookup"><span data-stu-id="42e15-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="42e15-105">选择获取最新版</span><span class="sxs-lookup"><span data-stu-id="42e15-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="42e15-106">选择获取最新版 Office：</span><span class="sxs-lookup"><span data-stu-id="42e15-106">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="42e15-107">如果您是 Office 365 家庭版、个人版或大专院校版订阅者，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider)。</span><span class="sxs-lookup"><span data-stu-id="42e15-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="42e15-108">如果您是 Office 365 商业版客户，请参阅 [为 Office 365 商业版客户安装首次发布](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)。</span><span class="sxs-lookup"><span data-stu-id="42e15-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="42e15-109">如果在 Mac 上运行 Office：</span><span class="sxs-lookup"><span data-stu-id="42e15-109">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="42e15-110">启动 Office for Mac 程序。</span><span class="sxs-lookup"><span data-stu-id="42e15-110">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="42e15-111">选择帮助菜单上的**检查更新**。</span><span class="sxs-lookup"><span data-stu-id="42e15-111">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="42e15-112">选中“Microsoft 自动更新”框，以加入 Office 预览体验成员计划。</span><span class="sxs-lookup"><span data-stu-id="42e15-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="42e15-113">获取最新版</span><span class="sxs-lookup"><span data-stu-id="42e15-113">Get the latest build</span></span>

<span data-ttu-id="42e15-114">获取最新版 Office：</span><span class="sxs-lookup"><span data-stu-id="42e15-114">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="42e15-115">下载  [Office 部署工具](https://www.microsoft.com/download/details.aspx?id=49117)。</span><span class="sxs-lookup"><span data-stu-id="42e15-115">Download the Office Deployment Tool</span></span> 
2. <span data-ttu-id="42e15-p101">运行该工具。这会提取以下两个文件：Setup.exe 和 configuration.xml。</span><span class="sxs-lookup"><span data-stu-id="42e15-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="42e15-118">将 configuration.xml 文件替换为[首次发布配置文件](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)。</span><span class="sxs-lookup"><span data-stu-id="42e15-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="42e15-119">以管理员身份运行下面的命令：  `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="42e15-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="42e15-120">此命令可能需要运行很长时间才能完成，而且不会显示进度。</span><span class="sxs-lookup"><span data-stu-id="42e15-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="42e15-121">在安装流程完成后，您已安装最新的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="42e15-121">When the installation process finishes, you will have the latest Office 2016 applications installed.</span></span> <span data-ttu-id="42e15-122">要验证你是否拥有最新版本，请从任何 Office 应用程序转到“**文件**” > “**帐户**”。</span><span class="sxs-lookup"><span data-stu-id="42e15-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="42e15-123">在Office 更新下，您将看到版本号上面的 (Office Insiders) 标签。</span><span class="sxs-lookup"><span data-stu-id="42e15-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![利用Office Insider标签显示产品信息的屏幕截图](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="42e15-125">Office JavaScript API 要求集对应的最低 Office 内部版本</span><span class="sxs-lookup"><span data-stu-id="42e15-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="42e15-126">若要了解 API 要求集对应的各个平台的最低产品内部版本，请参阅以下资源：</span><span class="sxs-lookup"><span data-stu-id="42e15-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="42e15-127">Word JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="42e15-127">Word JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets?view=office-js)
- [<span data-ttu-id="42e15-128">Excel JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="42e15-128">Excel JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js)
- [<span data-ttu-id="42e15-129">OneNote JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="42e15-129">OneNote JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets?view=office-js)
- [<span data-ttu-id="42e15-130">对话框 API 要求集</span><span class="sxs-lookup"><span data-stu-id="42e15-130">Dialog API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [<span data-ttu-id="42e15-131">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="42e15-131">Office common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js)

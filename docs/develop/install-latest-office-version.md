---
title: 安装最新版本 Office
description: 与如何选择获取最新版 Office 相关的信息。
ms.date: 01/14/2020
localization_priority: Priority
ms.openlocfilehash: 8d248e17433a2b6b1b31ad576ecdd237cd8543d0
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217379"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="024d8-103">安装最新版本 Office</span><span class="sxs-lookup"><span data-stu-id="024d8-103">Install the latest version of Office</span></span>

<span data-ttu-id="024d8-104">新开发人员功能（包括仍处于预览阶段的功能）会先向选择获取最新版 Office 的订阅者提供。</span><span class="sxs-lookup"><span data-stu-id="024d8-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span>

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="024d8-105">选择获取最新版</span><span class="sxs-lookup"><span data-stu-id="024d8-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="024d8-106">若要选择获取最新版 Office，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="024d8-106">To opt in to getting the latest builds of Office:</span></span>

- <span data-ttu-id="024d8-107">如果你是 Office 365 家庭版、个人版或大专院校版订阅者，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider)。</span><span class="sxs-lookup"><span data-stu-id="024d8-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="024d8-108">如果你是 Office 365 商业版客户，请参阅 [为 Office 365 商业版客户安装首次发布](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)。</span><span class="sxs-lookup"><span data-stu-id="024d8-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="024d8-109">如果在 Mac 上运行 Office：</span><span class="sxs-lookup"><span data-stu-id="024d8-109">If you're running Office on a Mac:</span></span>
  - <span data-ttu-id="024d8-110">启动 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="024d8-110">Start an Office application.</span></span>
  - <span data-ttu-id="024d8-111">选择“帮助”菜单上的“检查更新”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="024d8-111">Select **Check for Updates** on the Help menu.</span></span>
  - <span data-ttu-id="024d8-112">选中“Microsoft 自动更新”框，以加入 Office 预览体验成员计划。</span><span class="sxs-lookup"><span data-stu-id="024d8-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span>

## <a name="get-the-latest-build"></a><span data-ttu-id="024d8-113">获取最新版</span><span class="sxs-lookup"><span data-stu-id="024d8-113">Get the latest build</span></span>

<span data-ttu-id="024d8-114">若要获取最新版 Office，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="024d8-114">To get the latest build of Office:</span></span>

1. <span data-ttu-id="024d8-115">下载 [Office 部署工具](https://www.microsoft.com/download/details.aspx?id=49117)。</span><span class="sxs-lookup"><span data-stu-id="024d8-115">Download the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span>
2. <span data-ttu-id="024d8-p101">运行该工具。这会提取以下两个文件：Setup.exe 和 configuration.xml。</span><span class="sxs-lookup"><span data-stu-id="024d8-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="024d8-118">将 configuration.xml 文件替换为[首次发布配置文件](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)。</span><span class="sxs-lookup"><span data-stu-id="024d8-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="024d8-119">以管理员身份运行下面的命令：`setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="024d8-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span>

> [!NOTE]
> <span data-ttu-id="024d8-120">此命令可能需要运行很长时间才能完成，而且不会显示进度。</span><span class="sxs-lookup"><span data-stu-id="024d8-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="024d8-121">在安装进程完成后，你已安装最新的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="024d8-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="024d8-122">要验证你是否拥有最新版本，请从任何 Office 应用程序转到“文件”\*\*\*\* > “帐户”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="024d8-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="024d8-123">在“Office 更新”下，你将看到版本号上面的 (Office Insiders) 标签。</span><span class="sxs-lookup"><span data-stu-id="024d8-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![显示产品信息的屏幕截图（带有 Office Insiders 标签）](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="024d8-125">Office JavaScript API 要求集对应的最低 Office 内部版本</span><span class="sxs-lookup"><span data-stu-id="024d8-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="024d8-126">若要了解 API 要求集对应的各个平台的最低产品内部版本，请参阅以下资源：</span><span class="sxs-lookup"><span data-stu-id="024d8-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="024d8-127">Excel JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="024d8-127">Excel JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)
- [<span data-ttu-id="024d8-128">OneNote JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="024d8-128">OneNote JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets)
- [<span data-ttu-id="024d8-129">Outlook JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="024d8-129">Outlook JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
- [<span data-ttu-id="024d8-130">PowerPoint JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="024d8-130">PowerPoint JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets)
- [<span data-ttu-id="024d8-131">Word JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="024d8-131">Word JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets)
- [<span data-ttu-id="024d8-132">对话框 API 要求集</span><span class="sxs-lookup"><span data-stu-id="024d8-132">Dialog API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="024d8-133">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="024d8-133">Office Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)

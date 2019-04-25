---
title: 使用脚本实验室浏览 Office JavaScript API
description: 使用脚本实验室浏览 Office JS API 并建立原型功能。
ms.topic: article
ms.date: 04/23/2019
localization_priority: Normal
ms.openlocfilehash: 76888716cec8bd1754b7baa22dfcfbe5af984ea5
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32640280"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="6f4ad-103">使用脚本实验室浏览 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="6f4ad-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="6f4ad-104">通过 office 应用商店免费提供的[脚本实验室外接程序](https://store.office.com/app.aspx?assetid=WA104380862), 您可以在使用 office 程序 (如 Excel 或 Word) 时浏览 office JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-104">The [Script Lab add-in](https://store.office.com/app.aspx?assetid=WA104380862), which is available free from the Office store, enables you to explore the Office JavaScript API while you are working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="6f4ad-105">当您在外接程序中原型和验证所需功能时, 脚本实验室是一个方便的工具, 可将其添加到开发工具包中。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="6f4ad-106">什么是脚本实验室？</span><span class="sxs-lookup"><span data-stu-id="6f4ad-106">What is Script Lab?</span></span>

<span data-ttu-id="6f4ad-107">脚本实验室是任何希望了解如何使用 Excel、Word 或 PowerPoint 中的 office JavaScript API 开发 office 外接程序的工具。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="6f4ad-108">它提供了智能感知功能, 以便您可以查看在摩纳哥框架 (由 Visual Studio Code 使用的相同框架) 中构建的可用功能。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="6f4ad-109">通过脚本实验室, 您可以访问示例库以快速试用功能, 也可以选择示例作为自己的代码的基础。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-109">Through Script Lab, you can access a library of samples to quickly try out features or you can choose a sample as the base for your own code.</span></span> <span data-ttu-id="6f4ad-110">此外, 您还可以通过向[office js](https://github.com/OfficeDev/office-js-snippets#office-js-snippets)存储库添加代码段来扩展示例库。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-110">You are also welcome to expand the sample library by adding snippets to the [office-js-snippets repo](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span></span> <span data-ttu-id="6f4ad-111">脚本实验室的另一个激动人心的功能是 beta 或 preview 功能, 如可以尝试使用[自定义函数](/office/dev/add-ins/excel/custom-functions-overview)。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-111">Another exciting feature of Script Lab is beta or preview functionality like [custom functions](/office/dev/add-ins/excel/custom-functions-overview) is available for you to try.</span></span>

> [!TIP]
> <span data-ttu-id="6f4ad-112">若要参与 beta 或 preview, 您可能需要注册[Office 预览体验成员计划](https://products.office.com/office-insider)。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-112">To participate in beta or preview, you may have to sign up for the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="6f4ad-113">我到目前为止听起来正常吗？</span><span class="sxs-lookup"><span data-stu-id="6f4ad-113">Sounds good so far?</span></span> <span data-ttu-id="6f4ad-114">查看此一分钟视频可查看脚本实验室的实际效果。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-114">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="6f4ad-115">[![显示在 Excel、Word 和 PowerPoint Online 中运行的脚本实验室的预览视频。](../images/screenshot-wide-youtube.png '脚本实验室预览视频')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="6f4ad-115">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint Online.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="script-lab-supported-clients"></a><span data-ttu-id="6f4ad-116">脚本实验室支持的客户端</span><span class="sxs-lookup"><span data-stu-id="6f4ad-116">Script Lab supported clients</span></span>

<span data-ttu-id="6f4ad-117">以下客户端上的 Excel、Word 和 PowerPoint 支持脚本实验室。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-117">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="6f4ad-118">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="6f4ad-118">Office 365 for Windows</span></span>
- <span data-ttu-id="6f4ad-119">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="6f4ad-119">Office 365 for Mac</span></span>
- <span data-ttu-id="6f4ad-120">Office Online</span><span class="sxs-lookup"><span data-stu-id="6f4ad-120">Office Online</span></span>
- <span data-ttu-id="6f4ad-121">适用于 Windows 的 Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="6f4ad-121">Office 2013 or later for Windows</span></span>
- <span data-ttu-id="6f4ad-122">适用于 Mac 的 Office 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="6f4ad-122">Office 2016 or later for Mac</span></span>

## <a name="next-steps"></a><span data-ttu-id="6f4ad-123">后续步骤</span><span class="sxs-lookup"><span data-stu-id="6f4ad-123">Next steps</span></span>

<span data-ttu-id="6f4ad-124">当您准备好创建 Office 加载项时, 请参阅首选 Office 应用程序的[5 分钟快速入门](/office/dev/add-ins/#5-minute-quick-starts)。</span><span class="sxs-lookup"><span data-stu-id="6f4ad-124">When you're ready to create your Office Add-in, see the [5-minute quick start](/office/dev/add-ins/#5-minute-quick-starts) for your preferred Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="6f4ad-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6f4ad-125">See also</span></span>

- [<span data-ttu-id="6f4ad-126">获取脚本实验室</span><span class="sxs-lookup"><span data-stu-id="6f4ad-126">Get Script Lab</span></span>](https://store.office.com/app.aspx?assetid=WA104380862)
- [<span data-ttu-id="6f4ad-127">了解有关脚本实验室的详细信息</span><span class="sxs-lookup"><span data-stu-id="6f4ad-127">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="6f4ad-128">注册开发计划</span><span class="sxs-lookup"><span data-stu-id="6f4ad-128">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)

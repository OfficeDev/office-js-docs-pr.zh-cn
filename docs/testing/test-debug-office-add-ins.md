---
title: 测试和调试 Office 加载项
description: 了解如何测试和调试 Office 加载项
ms.date: 05/19/2021
localization_priority: Priority
ms.openlocfilehash: f794225d5ece20a430b967c8aa81ea165b573e52
ms.sourcegitcommit: 0d3bf72f8ddd1b287bf95f832b7ecb9d9fa62a24
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/02/2021
ms.locfileid: "52727925"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="51969-103">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="51969-103">Test and debug Office Add-ins</span></span>

<span data-ttu-id="51969-104">本文包含有关测试、调试和排查 Office 加载项问题的指南。</span><span class="sxs-lookup"><span data-stu-id="51969-104">This article contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a><span data-ttu-id="51969-105">测试跨平台及多个版本的 Office</span><span class="sxs-lookup"><span data-stu-id="51969-105">Test cross-platform and for multiple versions of Office</span></span>

<span data-ttu-id="51969-106">Office 加载项跨主要平台运行，因此需要在用户可能运行 Office 的所有平台上测试加载项。</span><span class="sxs-lookup"><span data-stu-id="51969-106">Office Add-ins run across major platforms, so you need to test an add-in in all the platforms where your users might be running Office.</span></span> <span data-ttu-id="51969-107">这通常包括 Office 网页版、Windows 版 Office（包括订阅和一次购买）、Mac 版 Office、iOS 版 Office 和 Android 版 Office（适用于 Outlook 加载项）。</span><span class="sxs-lookup"><span data-stu-id="51969-107">This usually includes Office on the web, Office on Windows (both subscription and one-time purchase), Office on Mac, Office on iOS, and (for Outlook add-ins) Office on Android.</span></span> <span data-ttu-id="51969-108">但是，有些情况下，你可以确定你的任何用户都没有在某些平台上工作。</span><span class="sxs-lookup"><span data-stu-id="51969-108">However, there may be some situations in which you can be sure that none of your users will be working on some platforms.</span></span> <span data-ttu-id="51969-109">例如，如果你正在为一家公司创建加载项，该公司要求其用户使用 Windows 计算机和订阅 Office，则无需针对 Office on Mac 或 一次性购买的 Windows 进行测试。</span><span class="sxs-lookup"><span data-stu-id="51969-109">For example, if you are making an add-in for a company that requires its users to work with Windows computers and subscription Office, then you don't need to test for Office on Mac or one-time purchase Windows.</span></span> 

> [!NOTE]
> <span data-ttu-id="51969-110">在 Windows 计算机上，Windows 和 Office 版本将决定加载项使用哪个浏览器控件。有关详细信息，请参阅 [加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="51969-110">On Windows computers, the version of Windows and Office will determine which browser control is used by add-ins. For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="51969-111">通过 AppSource 营销的加载项通过了包括在所有平台上测试的验证过程。</span><span class="sxs-lookup"><span data-stu-id="51969-111">Add-ins marketed through AppSource go through a validation process that includes testing on all platforms.</span></span> <span data-ttu-id="51969-112">此外，加载项已通过所有主要新式浏览器（包括 Microsoft Edge（基于 Chromium 的 WebView2）、Chrome 和 Safari）针对 Office 网页版进行了测试。</span><span class="sxs-lookup"><span data-stu-id="51969-112">In addition, add-ins are tested for Office on the web with all major modern browsers, including Microsoft Edge (Chromium-based WebView2), Chrome, and Safari.</span></span> <span data-ttu-id="51969-113">因此，提交 AppSource 之前，应在这些平台和浏览器上先进行测试。</span><span class="sxs-lookup"><span data-stu-id="51969-113">Accordingly, you should test on these platforms and browsers before you submit to AppSource.</span></span> <span data-ttu-id="51969-114">有关验证详细信息，请参阅 [商业市场证书策略](/legal/marketplace/certification-policies)，尤其是 [第 1120.3 一节](/legal/marketplace/certification-policies#11203-functionality)，以及 [Office 加载项应用程序和可用性页面](../overview/office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="51969-114">For more information about validation, see [Commercial marketplace certification policies](/legal/marketplace/certification-policies), especially [section 1120.3](/legal/marketplace/certification-policies#11203-functionality), and the [Office Add-in application and availability page](../overview/office-add-in-availability.md).</span></span> 
>
> <span data-ttu-id="51969-115">AppSource 不使用 Internet Explorer 或旧版 Microsoft Edge (WebView1) 测试 Office 网页版中的加载项。</span><span class="sxs-lookup"><span data-stu-id="51969-115">AppSource does not use Internet Explorer or the legacy version of Microsoft Edge (WebView1) to test add-ins in Office on the web.</span></span> <span data-ttu-id="51969-116">但如果有大量用户使用这两种浏览器打开 Office 网页版，则应使用这两种浏览器进行测试。</span><span class="sxs-lookup"><span data-stu-id="51969-116">But if a significant number of your users will use these two browsers to open Office on the web, then you should test with them.</span></span> <span data-ttu-id="51969-117">有关详细信息，请参阅 [支持 Internet Explorer 11](../develop/support-ie-11.md) 和 [Microsoft Edge 问题疑难解答](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。</span><span class="sxs-lookup"><span data-stu-id="51969-117">For more information, see [Support Internet Explorer 11](../develop/support-ie-11.md) and [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span></span> <span data-ttu-id="51969-118">Office 仍然支持这些浏览器的加载项，因此，如果认为你在加载项在这些浏览器中的运行方式方面遇到 bug，请为 [office-js](https://github.com/OfficeDev/office-js/issues/new/choose) 存储库创建问题。</span><span class="sxs-lookup"><span data-stu-id="51969-118">Office still supports these browsers for add-ins, so if you think you've encountered a bug in how add-ins run in them, please create an issue for the [office-js](https://github.com/OfficeDev/office-js/issues/new/choose) repo.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="51969-119">旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="51969-119">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="51969-p104">可以通过旁加载来安装 Office 加载项以供测试，而无需先将它添加到加载项目录中。 加载项的旁加载过程因平台而异，在某些情况下，也因产品而异。 下面的文章分别介绍了如何在特定平台或产品中旁加载 Office 加载项：</span><span class="sxs-lookup"><span data-stu-id="51969-p104">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product:</span></span>

- [<span data-ttu-id="51969-123">在 Windows 上旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="51969-123">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="51969-124">在 Office 网页版中旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="51969-124">Sideload Office Add-ins in Office on the web</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="51969-125">在 iPad 和 Mac 上旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="51969-125">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- [<span data-ttu-id="51969-126">旁加载 Outlook 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="51969-126">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a><span data-ttu-id="51969-127">调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="51969-127">Debug an Office Add-in</span></span>

<span data-ttu-id="51969-p105">Office 加载项的调试过程也因平台而异。 下面的文章分别介绍了如何在特定平台上调试 Office 加载项：</span><span class="sxs-lookup"><span data-stu-id="51969-p105">The procedure for debugging an Office Add-in varies by platform as well. Each of the following articles describes how to debug Office Add-ins on a specific platform:</span></span>

- [<span data-ttu-id="51969-130">从任务窗格附加调试器（在 Windows 上）</span><span class="sxs-lookup"><span data-stu-id="51969-130">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="51969-131">在 Windows 10 上使用 F12 开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="51969-131">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="51969-132">在 Office 网页版中调试加载项</span><span class="sxs-lookup"><span data-stu-id="51969-132">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="51969-133">在 iPad 和 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="51969-133">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

- [<span data-ttu-id="51969-134">适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展</span><span class="sxs-lookup"><span data-stu-id="51969-134">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="51969-135">验证 Office 加载项清单</span><span class="sxs-lookup"><span data-stu-id="51969-135">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="51969-136">若要了解如何验证描述 Office 加载项的清单文件，以及如何排查清单文件问题，请参阅[验证并排查清单问题](troubleshoot-manifest.md)。</span><span class="sxs-lookup"><span data-stu-id="51969-136">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="51969-137">排查用户错误</span><span class="sxs-lookup"><span data-stu-id="51969-137">Troubleshoot user errors</span></span>

<span data-ttu-id="51969-138">若要了解如何解决用户在使用 Office 加载项时可能会遇到的常见问题，请参阅[排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)。</span><span class="sxs-lookup"><span data-stu-id="51969-138">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>

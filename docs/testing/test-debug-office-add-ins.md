---
title: 测试和调试 Office 加载项
description: ''
ms.date: 11/24/2017
localization_priority: Priority
ms.openlocfilehash: 7ffa281807ca1541f8ebcc5f722c1043db115509
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388834"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="6dbdc-102">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6dbdc-102">Test and debug Office Add-ins</span></span>

<span data-ttu-id="6dbdc-103">本部分介绍了如何测试、调试和排查 Office 加载项问题。</span><span class="sxs-lookup"><span data-stu-id="6dbdc-103">This section contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="6dbdc-104">旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="6dbdc-104">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="6dbdc-105">可以通过旁加载来安装 Office 加载项以供测试，而无需先将它添加到加载项目录中。</span><span class="sxs-lookup"><span data-stu-id="6dbdc-105">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog.</span></span> <span data-ttu-id="6dbdc-106">加载项的旁加载过程因平台而异，在某些情况下，也因产品而异。</span><span class="sxs-lookup"><span data-stu-id="6dbdc-106">The procedure for sideloading an add-in varies by platform, and in some cases, by product as well.</span></span> <span data-ttu-id="6dbdc-107">下面的文章分别介绍了如何在特定平台或产品中旁加载 Office 加载项：</span><span class="sxs-lookup"><span data-stu-id="6dbdc-107">The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product:</span></span>

- [<span data-ttu-id="6dbdc-108">在 Windows 上旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6dbdc-108">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="6dbdc-109">在 Office Online 中旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6dbdc-109">Sideload Office Add-ins in Office Online</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="6dbdc-110">在 iPad 和 Mac 上旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6dbdc-110">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- [<span data-ttu-id="6dbdc-111">旁加载 Outlook 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="6dbdc-111">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="debug-an-office-add-in"></a><span data-ttu-id="6dbdc-112">调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6dbdc-112">Debug an Office Add-in</span></span>

<span data-ttu-id="6dbdc-113">Office 加载项的调试过程也因平台而异。</span><span class="sxs-lookup"><span data-stu-id="6dbdc-113">The procedure for debugging an Office Add-in varies by platform as well.</span></span> <span data-ttu-id="6dbdc-114">下面的文章分别介绍了如何在特定平台上调试 Office 加载项：</span><span class="sxs-lookup"><span data-stu-id="6dbdc-114">Each of the following articles describes how to debug Office Add-ins on a specific platform:</span></span>

- [<span data-ttu-id="6dbdc-115">从任务窗格附加调试器（在 Windows 上）</span><span class="sxs-lookup"><span data-stu-id="6dbdc-115">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="6dbdc-116">在 Windows 10 上使用 F12 开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="6dbdc-116">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="6dbdc-117">在 Office Online 中调试加载项</span><span class="sxs-lookup"><span data-stu-id="6dbdc-117">Debug add-ins in Office Online</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="6dbdc-118">在 iPad 和 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6dbdc-118">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="6dbdc-119">验证 Office 加载项清单</span><span class="sxs-lookup"><span data-stu-id="6dbdc-119">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="6dbdc-120">若要了解如何验证描述 Office 加载项的清单文件，以及如何排查清单文件问题，请参阅[验证并排查清单问题](troubleshoot-manifest.md)。</span><span class="sxs-lookup"><span data-stu-id="6dbdc-120">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="6dbdc-121">排查用户错误</span><span class="sxs-lookup"><span data-stu-id="6dbdc-121">Troubleshoot user errors</span></span>

<span data-ttu-id="6dbdc-122">若要了解如何解决用户在使用 Office 加载项时可能会遇到的常见问题，请参阅[排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)。</span><span class="sxs-lookup"><span data-stu-id="6dbdc-122">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>

---
title: 测试和调试 Office 加载项
description: 了解如何测试和调试 Office 加载项
ms.date: 05/19/2021
localization_priority: Priority
ms.openlocfilehash: 5df42a6c22325528eaaf2dcde28fddbfd3a211fb
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349355"
---
# <a name="test-and-debug-office-add-ins"></a>测试和调试 Office 加载项

本文包含有关测试、调试和排查 Office 加载项问题的指南。

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a>测试跨平台及多个版本的 Office

Office 加载项跨主要平台运行，因此需要在用户可能运行 Office 的所有平台上测试加载项。 这通常包括 Office 网页版、Windows 版 Office（包括订阅和一次购买）、Mac 版 Office、iOS 版 Office 和 Android 版 Office（适用于 Outlook 加载项）。 但是，有些情况下，你可以确定你的任何用户都没有在某些平台上工作。 例如，如果你正在为一家公司创建加载项，该公司要求其用户使用 Windows 计算机和订阅 Office，则无需针对 Office on Mac 或 一次性购买的 Windows 进行测试。

> [!NOTE]
> 在 Windows 计算机上，Windows 和 Office 版本将决定加载项使用哪个浏览器控件。有关详细信息，请参阅 [加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!IMPORTANT]
> 通过 AppSource 营销的加载项通过了包括在所有平台上测试的验证过程。 此外，加载项已通过所有主要新式浏览器（包括 Microsoft Edge（基于 Chromium 的 WebView2）、Chrome 和 Safari）针对 Office 网页版进行了测试。 因此，提交 AppSource 之前，应在这些平台和浏览器上先进行测试。 有关验证详细信息，请参阅 [商业市场证书策略](/legal/marketplace/certification-policies)，尤其是 [第 1120.3 一节](/legal/marketplace/certification-policies#11203-functionality)，以及 [Office 加载项应用程序和可用性页面](../overview/office-add-in-availability.md)。
>
> AppSource 不使用 Internet Explorer 或旧版 Microsoft Edge (WebView1) 测试 Office 网页版中的加载项。 但如果有大量用户使用这两种浏览器打开 Office 网页版，则应使用这两种浏览器进行测试。 有关详细信息，请参阅 [支持 Internet Explorer 11](../develop/support-ie-11.md) 和 [Microsoft Edge 问题疑难解答](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。 Office 仍然支持这些浏览器的加载项，因此，如果认为你在加载项在这些浏览器中的运行方式方面遇到 bug，请为 [office-js](https://github.com/OfficeDev/office-js/issues/new/choose) 存储库创建问题。

## <a name="sideload-an-office-add-in-for-testing"></a>旁加载 Office 加载项以供测试

可以通过旁加载来安装 Office 加载项以供测试，而无需先将它添加到加载项目录中。 加载项的旁加载过程因平台而异，在某些情况下，也因产品而异。 下面的文章分别介绍了如何在特定平台或产品中旁加载 Office 加载项。

- [在 Windows 上旁加载 Office 加载项](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [在 Office 网页版中旁加载 Office 加载项](sideload-office-add-ins-for-testing.md)

- [在 iPad 和 Mac 上旁加载 Office 加载项](sideload-an-office-add-in-on-ipad-and-mac.md)

- [旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a>调试 Office 加载项

Office 加载项的调试过程也因平台而异。 下面的文章分别介绍了如何在特定平台上调试 Office 加载项。

- [从任务窗格附加调试器（在 Windows 上）](attach-debugger-from-task-pane.md)

- [在 Windows 10 上使用 F12 开发人员工具调试加载项](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [在 Office 网页版中调试加载项](debug-add-ins-in-office-online.md)

- [在 iPad 和 Mac 上调试 Office 加载项](debug-office-add-ins-on-ipad-and-mac.md)

- [适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a>验证 Office 加载项清单

若要了解如何验证描述 Office 加载项的清单文件，以及如何排查清单文件问题，请参阅[验证并排查清单问题](troubleshoot-manifest.md)。

## <a name="troubleshoot-user-errors"></a>排查用户错误

若要了解如何解决用户在使用 Office 加载项时可能会遇到的常见问题，请参阅[排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)。

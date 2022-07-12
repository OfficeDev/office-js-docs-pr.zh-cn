---
title: 测试 Office 加载项
description: 了解如何测试 Office 加载项。
ms.date: 12/02/2021
ms.localizationpriority: high
ms.openlocfilehash: d69d57e677e7f06457f49fef60df63bc6f9577fa
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712732"
---
# <a name="test-office-add-ins"></a>测试 Office 加载项

本文包含有关测试、调试和排查 Office 加载项问题的指南。

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a>测试跨平台及多个版本的 Office

Office 加载项跨主要平台运行，因此需要在用户可能运行 Office 的所有平台上测试加载项。 这通常包括 Office 网页版、Windows 版 Office（包括订阅和一次购买）、Mac 版 Office、iOS 版 Office 和 Android 版 Office（适用于 Outlook 加载项）。 但是，有些情况下，你可以确定你的任何用户都没有在某些平台上工作。 例如，如果你正在为一家公司创建加载项，该公司要求其用户使用 Windows 计算机和订阅 Office，则无需针对 Office on Mac 或 一次性购买的 Windows 进行测试。

> [!NOTE]
> 在 Windows 计算机上，Windows 和 Office 版本将决定加载项使用哪个浏览器控件。有关详细信息，请参阅 [加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!IMPORTANT]
> 通过 AppSource 营销的加载项通过了包括在所有平台上测试的验证过程。 此外，加载项已通过所有主要新式浏览器（包括 Microsoft Edge（基于 Chromium 的 WebView2）、Chrome 和 Safari）针对 Office 网页版进行了测试。 因此，提交 AppSource 之前，应在这些平台和浏览器上先进行测试。 有关验证详细信息，请参阅 [商业市场证书策略](/legal/marketplace/certification-policies)，尤其是 [第 1120.3 一节](/legal/marketplace/certification-policies#11203-functionality)，以及 [Office 加载项应用程序和可用性页面](/javascript/api/requirement-sets)。
>
> AppSource 不使用 Internet Explorer 或旧版 Microsoft Edge (WebView1) 测试 Office 网页版中的加载项。 但如果有大量用户使用旧版 Edge 在 Web 上打开 Office，则需要进行测试。 (Office 网页版无法在 Internet Explorer 中打开，因此你无法也不需要使用 Internet Explorer 在 Web 上测试 Office。) 有关详细信息，请参阅[支持Internet Explorer 11](../develop/support-ie-11.md)和[Microsoft Edge 疑难解答](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。 Office 仍然支持这些浏览器的加载项，因此，如果您认为加载项在浏览器中运行时遇到 bug，请为[ office-js](https://github.com/OfficeDev/office-js/issues/new/choose) 存储库创建问题。

## <a name="sideload-an-office-add-in-for-testing"></a>旁加载 Office 加载项以供测试

可以通过旁加载来安装 Office 加载项以供测试，而无需先将它添加到加载项目录中。 加载项的旁加载过程因平台而异，在某些情况下，也因产品而异。 下面的文章分别介绍了如何在特定平台或产品中旁加载 Office 加载项。

- [在 Windows 上旁加载 Office 加载项](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [在 Office 网页版中旁加载 Office 加载项](sideload-office-add-ins-for-testing.md)

- [在 Mac 上旁加载 Office 加载项](sideload-an-office-add-in-on-mac.md)

- [在 iPad 上旁加载 Office 加载项](sideload-an-office-add-in-on-ipad.md)

- [旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="unit-testing"></a>单元测试

若要了解如何向加载项项目添加单元测试，请参阅[ Office 加载项单元测试](unit-testing.md)。

## <a name="debug-an-office-add-in"></a>调试 Office 加载项

调试 Office 加载项的过程因平台和环境而异。 有关详细信息，请参阅 [调试 Office 加载项](debug-add-ins-overview.md)。

## <a name="validate-an-office-add-in-manifest"></a>验证 Office 加载项清单

若要了解如何验证描述 Office 加载项的清单文件，以及如何排查清单文件问题，请参阅[验证并排查清单问题](troubleshoot-manifest.md)。

## <a name="troubleshoot-user-errors"></a>排查用户错误

若要了解如何解决用户在使用 Office 加载项时可能会遇到的常见问题，请参阅[排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)。

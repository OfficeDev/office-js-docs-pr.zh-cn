---
title: 测试和调试 Office 加载项
description: 了解如何测试和调试 Office 加载项
ms.date: 06/17/2020
localization_priority: Priority
ms.openlocfilehash: 526204fe94d4c97ce7e1e0bc9ac2a212f69611d3
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159246"
---
# <a name="test-and-debug-office-add-ins"></a>测试和调试 Office 加载项

本部分介绍了如何测试、调试和排查 Office 加载项问题。

## <a name="sideload-an-office-add-in-for-testing"></a>旁加载 Office 加载项以供测试

可以通过旁加载来安装 Office 加载项以供测试，而无需先将它添加到加载项目录中。 加载项的旁加载过程因平台而异，在某些情况下，也因产品而异。 下面的文章分别介绍了如何在特定平台或产品中旁加载 Office 加载项：

- [在 Windows 上旁加载 Office 加载项](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [在 Office 网页版中旁加载 Office 加载项](sideload-office-add-ins-for-testing.md)

- [在 iPad 和 Mac 上旁加载 Office 加载项](sideload-an-office-add-in-on-ipad-and-mac.md)

- [旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a>调试 Office 加载项

Office 加载项的调试过程也因平台而异。 下面的文章分别介绍了如何在特定平台上调试 Office 加载项：

- [从任务窗格附加调试器（在 Windows 上）](attach-debugger-from-task-pane.md)

- [在 Windows 10 上使用 F12 开发人员工具调试加载项](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [在 Office 网页版中调试加载项](debug-add-ins-in-office-online.md)

- [在 iPad 和 Mac 上调试 Office 加载项](debug-office-add-ins-on-ipad-and-mac.md)

- [适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a>验证 Office 加载项清单

若要了解如何验证描述 Office 加载项的清单文件，以及如何排查清单文件问题，请参阅[验证并排查清单问题](troubleshoot-manifest.md)。

## <a name="troubleshoot-user-errors"></a>排查用户错误

若要了解如何解决用户在使用 Office 加载项时可能会遇到的常见问题，请参阅[排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)。

---
title: 在 Office 网页版中调试加载项
description: 如何使用 Office 网页版来测试和调试加载项。
ms.date: 03/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5a07185c064d65432c7a3afce1e9f32e99034c3e
ms.sourcegitcommit: 3d7792b1f042db589edb74a895fcf6d7ced63903
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2022
ms.locfileid: "63435688"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>在 Office 网页版中调试加载项

本文介绍如何使用Office web 版调试外接程序。使用此技术：

- 在未运行 Windows 或 Office&mdash; 桌面客户端的计算机上调试外接程序，例如，如果你正在 Mac 或 Linux 上开发。
- 如果无法或不希望在 IDE 中调试（如在 IDE 中调试，Visual Studio调试Visual Studio Code。

本文假定您具有需要调试的外接程序项目。 如果只想在 Web 上练习调试，请为特定 Office 应用程序使用快速入门之一创建一个新项目，例如 [Word 的](../quickstarts/word-quickstart.md)此快速入门。

## <a name="debug-your-add-in"></a>调试加载项

若要使用 Office 网页版调试加载项，请执行以下操作：

1. 在 localhost 上运行项目，并旁加载它到Office web 版。 有关详细的旁加载说明，请参阅旁[加载Office Web 上的外接程序](sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web-manually)。

2. 打开浏览器的开发人员工具。 这通常通过按 F12 完成。 打开调试器工具并使用它来设置断点并监视变量。 有关使用浏览器工具的详细帮助，请参阅以下选项之一。  

   - [Firefox](https://developer.mozilla.org/en-US/docs/Tools)
   - [Safari](https://support.apple.com/guide/safari/use-the-developer-tools-in-the-develop-menu-sfri20948/mac)
   - [使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项](debug-add-ins-using-devtools-edge-chromium.md)
   - [使用旧版 Edge 开发人员工具调试加载项](debug-add-ins-using-devtools-edge-legacy.md)

   > [!NOTE]
   > Office web 版中打开Internet Explorer。

## <a name="potential-issues"></a>潜在问题

以下是在调试时可能会遇到的一些问题。

- 你看到的一些 JavaScript 错误可能源自 Office 网页版。

- 浏览器可能会显示无效证书错误，你需要忽略此错误。 执行此操作的过程因浏览器而异，而且用于执行此操作的各种浏览器的 UI 会定期进行更改。 有关说明，可搜索浏览器的“帮助”或“联机搜索”。 （例如，搜索“Microsoft Edge 无效证书警告”。）大多数浏览器在“警告”页面上都有一个链接，可以通过此链接单击进入“加载项”页。 例如，Microsoft Edge 有一个链接“转到网页（不推荐）”。 但是每次加载项重新加载时，通常都必须通过此链接来完成。 如需更长久地忽略，请参阅建议的帮助。

- 如果在代码中设置断点，Office web 版可能会抛出错误，指示无法保存。

## <a name="see-also"></a>另请参阅

- [Office 加载项开发最佳做法](../concepts/add-in-development-best-practices.md)
- [排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)

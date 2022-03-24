---
title: 使用 Visual Studio 开发 Office 加载项
description: 如何使用 Visual Studio 开发 Office 加载项。
ms.date: 01/26/2022
ms.localizationpriority: high
ms.openlocfilehash: 52740e16363e3e038269e08a9e50e0f08877db66
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743840"
---
# <a name="develop-office-add-ins-with-visual-studio"></a>使用 Visual Studio 开发 Office 加载项

本文介绍如何使用 Visual Studio 开发 Office 加载项。 如果你已创建加载项，则可以跳至[使用 Visual Studio 开发加载项](#develop-the-add-in-using-visual-studio)部分。

> [!NOTE]
> 作为使用 Visual Studio 的替代方法，可以选择使用 Office 加载项的 Yeoman 生成器和 VS Code 创建 Office 加载项。要了解关于此选项的详细信息，请参阅 [创建 Office 加载项](../develop/develop-overview.md#create-an-office-add-in)。

## <a name="create-the-add-in-project-using-visual-studio"></a>使用 Visual Studio 创建加载项项目

Visual Studio 可用于创建适用于 Excel、Outlook、Word 和 PowerPoint 的 Office 加载项。 Office 加载项项目是作为 Visual Studio 解决方案的一部分创建的，它使用 HTML、CSS 和 JavaScript。 要使用 Visual Studio 创建 Office 加载项，请按照快速入门中与要创建的加载项相对应的说明操作。  

- [Excel 快速入门](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=visualstudio)
- [Word 快速入门](../quickstarts/word-quickstart.md?tabs=visualstudio)
- [PowerPoint 快速入门](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)

Visual Studio 不支持为 OneNote 或 Project 创建 Office 加载项。如果要为其中任一应用程序创建 Office 加载项，需要使用 Office 加载项的 Yeoman 生成器，如 [OneNote 快速入门](../quickstarts/onenote-quickstart.md) 或 [Project 快速入门](../quickstarts/project-quickstart.md) 中所述。

## <a name="develop-the-add-in-using-visual-studio"></a>使用 Visual Studio 开发加载项

Visual Studio 会创建一个功能受限的基本加载项。 你可通过在 Visual Studio 中编辑[清单](add-in-manifests.md)、HTML、JavaScript 和 CSS 文件来自定义加载项。 有关 Visual Studio 创建的加载项项目中的项目结构和文件的高级说明，请参阅用于指导创建加载项的快速入门中的 Visual Studio 指南。

> [!TIP]
> 由于 Office 加载项是一种 Web 应用程序，因此你至少需要具备基本的 Web 开发技能才能自定义加载项。 如果你不熟悉 JavaScript，建议查看 [Mozilla JavaScript 教程](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)。

要自定义加载项，你需要了解本文档的“[核心概念 > 开发](develop-overview.md)”区域中描述的概念，以及与要构建的加载项相对应的文档应用程序特定区域中描述的概念（例如，[Excel](../excel/index.yml)）。

## <a name="test-and-debug-the-add-in"></a>测试和调试加载项

用于测试、调试和故障排除 Office 加载项的方法因平台而异。 有关详细信息，请参阅[在 Visual Studio 中调试 Office 加载项](debug-office-add-ins-in-visual-studio.md)和[测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)。

## <a name="publish-the-add-in"></a>发布加载项

Office 加载项 包含 Web 应用程序和清单文件。Web 应用程序定义加载项的用户界面和功能，清单指定 Web 应用程序的位置并定义加载项的设置和功能。

在 Visual Studio 中开发加载项时，该加载项将在本地 Web 服务器 (`localhost`) 上运行。 如果加载项如期工作且你已准备好发布它供其他用户访问，你需要完成以下步骤。

1. 将 Web 应用程序部署到 Web 服务器或 Web 托管服务（例如 Microsoft Azure）。
2. 更新清单以指定已部署应用程序的 URL。
3. 选择要用来[部署 Office 加载项](../publish/publish.md)的方法，再按照说明发布清单文件。

## <a name="see-also"></a>另请参阅

- [Office 加载项的核心概念](../overview/core-concepts-office-add-ins.md)
- [开发 Office 加载项](../develop/develop-overview.md)
- [设计 Office 加载项](../design/add-in-design.md)
- [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)
- [发布 Office 加载项](../publish/publish.md)

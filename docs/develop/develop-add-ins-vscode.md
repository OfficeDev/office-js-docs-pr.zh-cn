---
title: 使用 Visual Studio Code 开发 Office 加载项
description: 如何使用 Visual Studio Code 开发 Office 加载项。
ms.date: 02/18/2022
ms.localizationpriority: high
ms.openlocfilehash: 6710884a9bc751e6a94607581223dabaea0bce3b
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63511292"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a>使用 Visual Studio Code 开发 Office 加载项

本文介绍如何使用 [Visual Studio Code (VS Code)](https://code.visualstudio.com) 开发 Office 加载项。

> [!NOTE]
> 要了解如何使用 Visual Studio 创建 Office 加载项，请参阅[使用 Visual Studio 开发 Office 加载项](develop-add-ins-visual-studio.md)。

## <a name="prerequisites"></a>先决条件

- [Visual Studio Code](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a>使用 Yeoman 生成器创建加载项项目

如果你正在将 VS Code 用作集成开发环境 (IDE)，则应使用[适用于 Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)来创建 Office 加载项项目。Yeoman 生成器会创建一个 Node.js 项目，它可通过 VS Code 或任何其他编辑器进行管理。

要使用 Yeoman 生成器创建 Office 加载项，请按照 [5 分钟快速入门](../index.yml)中与你要创建的加载项类型相对应的说明进行操作。

## <a name="develop-the-add-in-using-vs-code"></a>使用 VS Code 开发加载项

在 Yeoman 生成器完成加载项项目的创建后，请使用 VS Code 打开项目的根文件夹。

[!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

Yeoman 生成器会创建一个功能受限的基本加载项。 你可通过在 VS Code 中编辑[清单](add-in-manifests.md)HTML、JavaScript/TypeScript 和 CSS 文件，自定义该加载项。 要简要了解 Yeoman 生成器创建的加载项项目中的项目结构和文件，请查看 [5 分钟快速入门](../index.yml)中与你创建的加载项类型相对应的 Yeoman 生成器指南。

## <a name="test-and-debug-the-add-in"></a>测试和调试加载项

用于测试、调试和故障排除 Office 加载项的方法因平台而异。 有关详细信息，请参阅 [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)。

## <a name="publish-the-add-in"></a>发布加载项

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a>另请参阅

- [Office 加载项的核心概念](../overview/core-concepts-office-add-ins.md)
- [开发 Office 加载项](../develop/develop-overview.md)
- [设计 Office 加载项](../design/add-in-design.md)
- [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)
- [发布 Office 加载项](../publish/publish.md)
---
title: 设置开发环境
description: 设置开发人员环境以生成 Office 外接程序
ms.date: 04/03/2020
localization_priority: Normal
ms.openlocfilehash: f44f8e48aec402f0ffa6327732613a902ea0cfe6
ms.sourcegitcommit: 19312a54f47a17988ffa86359218a504713f9f09
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/10/2020
ms.locfileid: "44679351"
---
# <a name="set-up-your-development-environment"></a>设置开发环境

本指南可帮助您设置工具，以便您可以按照快速入门或教程创建 Office 加载项。 你将需要安装以下列表中的工具。 如果已安装了这些安装，则可以开始快速启动，例如此 Excel 会对[快速启动做出反应](../quickstarts/excel-quickstart-react.md)。

- Node.js
- npm
- Office 365 （Office 的订阅版本）帐户
- 您选择的代码编辑器

本指南假定您知道如何使用命令行工具。 

## <a name="install-nodejs"></a>安装 node.js

Node.js 是开发新式 Office 外接程序所需的 JavaScript 运行时。

通过[从网站下载最新的推荐版本](https://nodejs.org)来安装 node.js。 按照操作系统的安装说明进行操作。

## <a name="install-npm"></a>安装 npm

npm 是一个开放源代码软件注册表，可从中下载用于开发 Office 外接程序的程序包。

若要安装 npm，请在命令行中运行以下命令：

```command&nbsp;line
    npm install npm -g
```

若要检查是否已安装了 npm 并查看已安装的版本，请在命令行中运行以下命令：

```command&nbsp;line
npm -v
```

您可能希望使用节点版本管理器，以允许在多个版本的 node.js 和 npm 之间进行切换，但这并不是绝对必要的。 有关如何执行此操作的详细信息，[请参阅 npm 的说明](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。

## <a name="get-office-365"></a>获取 Office 365

如果还没有 Office 365 账户，可以通过加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获得 90 天免费的可续订 Office 365 订阅。

## <a name="install-a-code-editor"></a>安装代码编辑器

若要生成 Web 部件，可以使用任何支持客户端开发的代码编辑器或 IDE，如：

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>后续步骤

尝试创建您自己的外接程序，或使用脚本实验室来尝试内置的示例。

### <a name="create-an-office-add-in"></a>创建 Office 加载项

可完成 [5 分钟快速入门](/office/dev/add-ins/)，快速创建适合 Excel、OneNote、Outlook、PowerPoint、Project 或 Word 的基本加载项。 如果你之前已完成快速入门，并且想要创建更复杂一些的加载项，请尝试本[教程](/office/dev/add-ins/)。

### <a name="explore-the-apis-with-script-lab"></a>使用 Script Lab 了解 API

了解 [Script Lab](explore-with-script-lab.md) 中的内置示例库，熟悉 Office JavaScript API 的功能。

## <a name="see-also"></a>另请参阅

- [构建 Office 加载项](../overview/office-add-ins-fundamentals.md)
- [Office 加载项的核心概念](../overview/core-concepts-office-add-ins.md)
- [开发 Office 加载项](../develop/develop-overview.md)
- [设计 Office 加载项](../design/add-in-design.md)
- [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)
- [发布 Office 加载项](../publish/publish.md)

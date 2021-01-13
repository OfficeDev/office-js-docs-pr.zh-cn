---
title: 设置开发环境
description: 设置开发人员环境以生成 Office 加载项。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: eddf8bdf7b20a54667e6f8eb38bdace801ea1813
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839710"
---
# <a name="set-up-your-development-environment"></a>设置开发环境

本指南可帮助你设置工具，以便你可以按照我们的快速入门或教程创建 Office 外接程序。 你需要从下面的列表中安装工具。 如果已安装这些组件，则已准备好开始快速入门，例如 [此 Excel React 快速入门](../quickstarts/excel-quickstart-react.md)。

- Node.js
- npm
- 包含 Office 订阅版本的 Microsoft 365 帐户
- 你选择的代码编辑器

本指南假定您了解如何使用命令行工具。 

## <a name="install-nodejs"></a>安装 Node.js

Node.js JavaScript 运行时，你需要开发新式 Office 加载项。

通过Node.js下载最新建议版本 [来安装此应用](https://nodejs.org)。 按照操作系统的安装说明进行操作。

## <a name="install-npm"></a>安装 npm

npm 是一个开源软件注册表，可从中下载用于开发 Office 外接程序的程序包。

若要安装 npm，请运行命令行中的以下命令：

```command&nbsp;line
    npm install npm -g
```

若要检查是否已安装 npm 并查看已安装的版本，请在命令行中运行以下命令：

```command&nbsp;line
npm -v
```

你可能希望使用节点版本管理器来允许你在多个版本的 Node.js 和 npm 之间切换，但这不是严格必需的。 有关如何执行此操作的详细信息，请参阅 [npm 的说明](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。

## <a name="get-office-365"></a>获取 Office 365

如果还没有 Microsoft 365 帐户，可以通过加入 [Microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获得为期 90 天的免费可续订 Microsoft 365 订阅。

## <a name="install-a-code-editor"></a>安装代码编辑器

若要生成 Web 部件，可以使用任何支持客户端开发的代码编辑器或 IDE，如：

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>后续步骤

请尝试创建自己的外接程序或使用 Script Lab 试用内置示例。

### <a name="create-an-office-add-in"></a>创建 Office 加载项

可完成 [5 分钟快速入门](../index.yml)，快速创建适合 Excel、OneNote、Outlook、PowerPoint、Project 或 Word 的基本加载项。 如果你之前已完成快速入门，并且想要创建更复杂一些的加载项，请尝试本[教程](../index.yml)。

### <a name="explore-the-apis-with-script-lab"></a>使用 Script Lab 了解 API

了解 [Script Lab](explore-with-script-lab.md) 中的内置示例库，熟悉 Office JavaScript API 的功能。

## <a name="see-also"></a>另请参阅

- [Office 加载项的核心概念](../overview/core-concepts-office-add-ins.md)
- [开发 Office 加载项](../develop/develop-overview.md)
- [设计 Office 加载项](../design/add-in-design.md)
- [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)
- [发布 Office 加载项](../publish/publish.md)
- [了解 Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)
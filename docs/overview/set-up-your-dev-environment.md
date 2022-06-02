---
title: 设置开发环境
description: 设置开发人员环境以生成Office加载项。
ms.date: 05/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 01b9fe0aff2696a521266bb3175ea0f61d891aa4
ms.sourcegitcommit: 35e7646c5ad0d728b1b158c24654423d999e0775
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/02/2022
ms.locfileid: "65833883"
---
# <a name="set-up-your-development-environment"></a>设置开发环境

本指南可帮助你设置工具，以便你可以按照我们的快速入门或教程创建Office加载项。 如果已安装这些设备，则可以快速入门，如此[Excel React快速入门](../quickstarts/excel-quickstart-react.md)。

## <a name="get-microsoft-365"></a>获取Microsoft 365

你需要一个Microsoft 365帐户。 你可以通过加入Microsoft 365[开发人员计划](https://developer.microsoft.com/office/dev-program)获得包含所有Office应用的免费 90 天可再生Microsoft 365订阅。

## <a name="install-the-environment"></a>安装环境

有两种类型的开发环境可供选择。 在两个环境中创建的Office外接程序项目的基架不同，因此如果多人将处理加载项项目，则必须使用相同的环境。 

- **Node.js环境**：建议。 在此环境中，工具会安装并在命令行上运行。 外接程序的 Web 应用程序部件的服务器端以 JavaScript 或 TypeScript 编写，并托管在Node.js运行时中。 此环境中有许多有用的外接程序开发工具，例如Office linter 和名为 WebPack 的捆绑程序/任务运行程序。 项目创建和基架工具 Yo Office 会频繁更新。
- **Visual Studio环境**：仅当开发计算机Windows时才选择此环境，并且要使用基于 .NET 的语言和框架（例如 ASP.NET）开发加载项的服务器端。 Visual Studio中的外接程序项目模板的更新频率不如Node.js环境中的更新频率。 无法使用内置Visual Studio调试器调试客户端代码，但可以使用浏览器的开发工具调试客户端代码。 稍后有关 **Visual Studio环境** 选项卡的详细信息。

> [!NOTE]
> Visual Studio for Mac不包括用于Office加载项的项目基架模板，因此，如果开发计算机是 Mac，则应使用Node.js环境。

选择所选环境的选项卡。 

# <a name="nodejs-environment"></a>[Node.js环境](#tab/yeomangenerator)

要安装的主要工具包括：

- Node.js
- npm
- 所选代码编辑器
- Yo Office
- Office JavaScript linter

本指南假定你知道如何使用命令行工具。

### <a name="install-nodejs-and-npm"></a>安装Node.js和npm

Node.js是用于开发新式Office加载项的 JavaScript 运行时。

通过 [从其网站下载最新推荐版本来安装](https://nodejs.org)Node.js。 按照操作系统的安装说明操作。

npm是一个开放源代码软件注册表，从中下载用于开发Office加载项的包。安装Node.js时，通常会自动安装它。 若要检查是否已安装npm并查看已安装的版本，请在命令行中运行以下命令。

```command&nbsp;line
npm -v
```

如果出于任何原因想要手动安装它，请在命令行中运行以下命令。

```command&nbsp;line
npm install npm -g
```

> [!TIP]
> 你可能希望使用节点版本管理器来允许在多个版本的Node.js和npm之间切换，但这不是绝对必要的。 有关如何执行此操作的详细信息，请[参阅npm的说明](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。

### <a name="install-a-code-editor"></a>安装代码编辑器

若要生成 Web 部件，可以使用任何支持客户端开发的代码编辑器或 IDE，如：

- [建议Visual Studio Code](https://code.visualstudio.com/) () 
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

### <a name="install-the-yeoman-generator-mdash-yo-office"></a>安装 Yeoman 生成器 &mdash; Yo Office

项目创建和基架工具是 [用于Office外接程序的 Yeoman 生成器](../develop/yeoman-generator-overview.md)，亲切地称为 **Yo Office**。 需要安装最新版本的 [Yeoman](https://github.com/yeoman/yo) 和 Yo Office。 若要全局安装这些工具，请通过命令提示符运行以下命令。

  ```command&nbsp;line
  npm install -g yo generator-office
  ```

### <a name="install-and-use-the-office-javascript-linter"></a>安装并使用 Office JavaScript linter

Microsoft 提供了一个 JavaScript linter 来帮助你在使用 Office JavaScript 库时捕获常见错误。 若要安装 linter，请在[安装](#install-nodejs-and-npm)Node.js并npm) 后 (运行以下两个命令。

```command&nbsp;line
npm install office-addin-lint --save-dev
npm install eslint-plugin-office-addins --save-dev
```

如果使用 [Yeoman 生成器](../develop/yeoman-generator-overview.md)为Office加载项工具创建Office外接程序项目，则其余的设置将为你完成。 在编辑器的终端（如Visual Studio Code或命令提示符中）使用以下命令运行 linter。 Linter 发现的问题会显示在终端或提示符中，并且在使用支持 linter 消息的编辑器（例如Visual Studio Code）时也直接显示在代码中。  (有关安装 Yeoman 生成器的信息，请参阅[适用于 Office Add-ins.) 的 Yeoman 生成器](../develop/yeoman-generator-overview.md)

```command&nbsp;line
npm run lint
```

如果外接程序项目是通过另一种方式创建的，请执行以下步骤。

1. 在项目的根目录中，创建名为 **.eslintrc.json** 的文本文件（如果没有一个）。 请确保它具有命名 `plugins` 属性和 `extends`类型数组的属性。 该 `plugins` 数组应包括 `"office-addins"` 在内， `extends` 数组应包括在内 `"plugin:office-addins/recommended"`。 下面展示了一个非常简单的示例。 **.eslintrc.json** 文件可能具有两个数组的其他属性和其他成员。

   ```json
   {
     "plugins": [
       "office-addins"
     ],
     "extends": [
       "plugin:office-addins/recommended"
     ]
   }
   ```

1. 在项目的根目录中，打开 **package.json** 文件，并确保数 `scripts` 组具有以下成员。

   ```json
   "lint": "office-addin-lint check",
   ```

1. 在编辑器的终端（如Visual Studio Code或命令提示符中）使用以下命令运行 linter。 Linter 发现的问题会显示在终端或提示符中，并且在使用支持 linter 消息的编辑器（例如Visual Studio Code）时也直接显示在代码中。

   ```command&nbsp;line
   npm run lint
   ```

# <a name="visual-studio-environment"></a>[Visual Studio环境](#tab/visualstudio)

### <a name="install-visual-studio"></a>安装 Visual Studio

如果未安装 Visual Studio 2017 (Windows) 或更高版本，请从[Visual Studio下载](https://visualstudio.microsoft.com/downloads/)中安装最新版本。 当安装程序要求指定工作负荷时，请务必包含 **Office/SharePoint开发** 工作负荷。 可能需要的其他工作负载是用于 .NET、**JavaScript 和 TypeScript 语言的** **Web 开发工具**， (用于对外接程序) 的客户端进行编码，以及 ASP.NET 相关工作负荷。

> [!TIP]
> 从 2022 年夏季开始，随Visual Studio一起安装的 Office 外接程序清单的 XML 架构不是最新版本。 这可能会影响加载项，具体取决于它们使用的外接程序功能。 因此，可能需要更新清单的 XML 架构。 有关详细信息，请参阅[Visual Studio项目的清单架构验证错误](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)。

> [!NOTE]
> 有关使用Visual Studio环境时调试客户端代码的信息，请参阅[Visual Studio中的调试Office加载项](../develop/debug-office-add-ins-in-visual-studio.md)。 调试服务器端代码的方式与在Visual Studio中创建的任何 Web 应用程序相同。 请参阅 [客户端或服务器端](../testing/debug-add-ins-overview.md#server-side-or-client-side)。

---

## <a name="install-script-lab"></a>安装Script Lab

Script Lab是一种快速原型编写调用 Office JavaScript 库 API 的代码的工具。 Script Lab本身是一个Office加载项，可在 [Script Lab](https://appsource.microsoft.com/marketplace/apps?search=script%20lab&page=1) 从 AppSource 安装。 Excel、PowerPoint 和 Word 有一个版本，另外还有一个用于Outlook的单独版本。 有关如何使用Script Lab的信息，请[参阅使用Script Lab浏览Office JavaScript API](explore-with-script-lab.md)。

## <a name="next-steps"></a>后续步骤

尝试创建自己的加载项或使用[Script Lab](explore-with-script-lab.md)来尝试内置示例。

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
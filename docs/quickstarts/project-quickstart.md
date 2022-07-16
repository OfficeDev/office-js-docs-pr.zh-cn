---
title: 生成首个 Project 任务窗格加载项
description: 了解如何使用 Office JS API 生成简单的 Project 任务窗格加载项。
ms.date: 07/13/2022
ms.prod: project
ms.localizationpriority: high
ms.openlocfilehash: c2f0e31b5a4c958cd155dfeb6d1648f7a2697c69
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797475"
---
# <a name="build-your-first-project-task-pane-add-in"></a>生成首个 Project 任务窗格加载项

本文将逐步介绍如何生成 Project 任务窗格加载项。

## <a name="prerequisites"></a>先决条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Windows 版 Project 2016 或更高版本

## <a name="create-the-add-in"></a>创建加载项

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **选择项目类型:** `Office Add-in Task Pane project`
- **选择脚本类型:** `Javascript`
- **要如何命名加载项?** `My Office Add-in`
- **要支持哪一个 Office 客户端应用程序?** `Project`

![命令行界面中 Yeoman 生成器的提示和回答。](../images/yo-office-project.png)

完成此向导后，生成器会创建项目，并安装支持的 Node 组件。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>浏览项目

使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。

- 项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。
- **./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。
- **./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。
- **./src/taskpane/taskpane.js** 文件包含用于加快任务窗格与 Office 客户端应用程序之间的交互的 Office JavaScript API 代码。 在本快速入门中，代码设置了项目所选任务的 `Name` 字段和 `Notes` 字段。

## <a name="try-it-out"></a>试用

1. 导航到项目的根文件夹。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. 启动本地 Web 服务器。

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    在项目的根目录中运行以下命令。 运行此命令时，本地 Web 服务器将启动。

    ```command&nbsp;line
    npm run dev-server
    ```

1. 在 Project 中，创建一个简单的项目计划。

1. 按照[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)中的说明，在 Project 中加载你的加载项。

1. 在项目中选择单个任务。

1. 在任务窗格的底部，选择“**运行**”链接以重命名所选任务并向所选任务添加备注。

    ![加载了任务窗格加载项的 Project 应用程序。](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>后续步骤

恭喜！已成功创建 Project 加载项！接下来，请详细了解 Project 加载项功能，并探索常见方案。

> [!div class="nextstepaction"]
> [Project 加载项](../project/project-add-ins.md)

## <a name="see-also"></a>另请参阅

- [开发 Office 加载项](../develop/develop-overview.md)
- [Office 加载项的核心概念](../overview/core-concepts-office-add-ins.md)
- [使用 Visual Studio Code 发布](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)

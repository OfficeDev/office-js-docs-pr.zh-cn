---
title: 使用代码和 Azure Visual Studio外接程序
description: 如何使用 Code 和 Azure Active Directory Visual Studio加载项
ms.date: 08/12/2020
localization_priority: Normal
ms.openlocfilehash: 3552e4eebacc84fc2b8e37782c97b4e03e96e508
ms.sourcegitcommit: 7faa0932b953a4983a80af70f49d116c3236d81a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/21/2020
ms.locfileid: "46845507"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>发布使用 Visual Studio Code 开发的加载项

本文介绍如何发布使用 Yeoman 生成器创建并使用 [Visual Studio Code (VS Code)](https://code.visualstudio.com) 或任何其他编辑器开发的 Office 加载项。

> [!NOTE]
> 要了解如何发布使用 Visual Studio 创建的 Office 加载项，请参阅[使用 Visual Studio 发布加载项](package-your-add-in-using-visual-studio.md)。

## <a name="publishing-an-add-in-for-other-users-to-access"></a>发布加载项供其他人用户访问

Office 加载项由一个 Web 应用程序和一个清单文件构成。 Web 应用程序定义加载项的用户界面和功能，清单指定 Web 应用程序的位置并定义加载项的设置和功能。

在开发过程中，你可以在本地 Web 服务器上运行该加载项， (具体 `localhost`) 。 当您准备好发布它供其他用户访问时，你将需要部署 Web 应用程序并更新清单以指定已部署应用程序的 URL。

如果外接程序可按需工作，则可以使用 Azure 存储扩展直接通过 Visual Studio Code 发布它。

## <a name="using-visual-studio-code-to-publish"></a>使用Visual Studio发布

>[!NOTE]
> 这些步骤仅适用于使用 Yeoman 生成器创建的项目。

1. 在 VS Code 代码编辑器的代码Visual Studio从其 (根文件夹中) 。
2. 从 VS Code 中的扩展视图，搜索 Azure 存储扩展并将其安装。
3. 安装完成后，会向活动栏中添加一个 Azure 图标。 选择它可访问扩展。 如果你的活动栏处于隐藏状态，你将无法访问扩展。 通过选择视图和显示活动 **栏>显示>栏**。
4. 在扩展中时，选择"登录 Azure"，登录 **Azure 帐户**。 如果你还没有 Azure 帐户，也可以选择"创建免费的 Azure 帐户 **"帐户。** 请按照提供的步骤设置帐户。
5. 登录到 Azure 帐户后，你将会看到你的 Azure 存储帐户的显示有扩展中。 如果还没有存储帐户，则需要使用"创建新存储帐户 **"选项创建一个存储帐户** 。 对存储帐户命名一个全局唯一名称，仅使用"a-z"和"0-9"。 请注意，默认情况下，这将创建一个同名的存储帐户和资源组。 它将自动将存储帐户放置于美国西部。 可以通过 Azure 帐户联机 [调整该操作](https://portal.azure.com/)。
6. 选择并 (右键单击) ，选择"配置**静态网站"。** 系统将要求您输入索引文档的名称和 404 的文档名称。 将索引文档的名称从默认更改为 `index.html` **`taskpane.html`** . 您可能还要更改 404 个文档名称，但不是必需的。
7. 现在选择 (静态网站) 右键单击，然后右键单击 **邮箱**。 从打开的浏览器窗口中复制网站 URL。
8. 在 VS Code 中，打开项目的清单文件 (`manifest.xml`) 并将对 localhost URL (如 `https://localhost:3000`) 已复制的 URL 的任何引用。 此终结点是您新建的存储帐户的静态网站 URL。 保存对清单文件所做的更改。
9. 打开命令行提示符并导航到加载项项目的根目录。 然后运行以下命令，为生产部署准备所有文件。

    ```command&nbsp;line
    npm run build
    ```

    生成完成后，加载项项目的根目录中的 **dist** 文件夹将包含要在后续步骤中部署的文件。

10. 若要部署，请选择文件资源管理器，选择并 (右键单击) **不同的文件夹**，然后选择"**部署到静态网站"。** 出现提示时，选择之前创建的存储帐户。

![部署到静态网站](../images/deploy-to-static-website.png)

11. 部署完成后，将显示 **"浏览至网站** "消息，你可以选择该消息来打开已部署应用代码的主要终结点。

## <a name="see-also"></a>另请参阅

- [使用 Visual Studio Code 开发 Office 加载项](../develop/develop-add-ins-vscode.md)
- [部署和发布 Office 外接程序](../publish/publish.md)

---
title: 使用 Visual Studio Code 和 Azure 发布加载项
description: 如何使用 Visual Studio Code 和 Azure Active Directory 发布加载项
ms.date: 08/19/2022
ms.custom: vscode-azure-extension-update-completed
ms.localizationpriority: medium
ms.openlocfilehash: 1c82d62e9f92453839084179d7ef9e0a8e2c8ca3
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464782"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>发布使用 Visual Studio Code 开发的加载项

本文介绍如何发布使用 Yeoman 生成器创建并使用 [Visual Studio Code (VS Code)](https://code.visualstudio.com) 或任何其他编辑器开发的 Office 加载项。

> [!NOTE]
> 要了解如何发布使用 Visual Studio 创建的 Office 加载项，请参阅[使用 Visual Studio 发布加载项](package-your-add-in-using-visual-studio.md)。

## <a name="publishing-an-add-in-for-other-users-to-access"></a>发布加载项供其他人用户访问

Office 加载项 包含 Web 应用程序和清单文件。Web 应用程序定义加载项的用户界面和功能，清单指定 Web 应用程序的位置并定义加载项的设置和功能。

开发时，可以在本地 Web 服务器上运行加载项 () `localhost` 。 准备好将其发布供其他用户访问时，需要部署 Web 应用程序并更新清单以指定已部署应用程序的 URL。

当加载项按需工作时，可以使用 Azure 存储扩展直接通过Visual Studio Code发布。

## <a name="using-visual-studio-code-to-publish"></a>使用 Visual Studio Code 发布

>[!NOTE]
> 这些步骤仅适用于使用 Yeoman 生成器创建的项目。

1. 在 VISUAL STUDIO CODE (VS Code) 中从其根文件夹打开项目。
2. 在 VS Code 中的“扩展”视图中，搜索 Azure 存储扩展并安装它。
3. 安装后，Azure 图标将添加到活动栏。 选择它以访问扩展。 如果活动栏处于隐藏状态，则无法访问扩展。 通过选择“ **查看>外观>活动栏”显示活动栏**。
4. 运行扩展并选择 **登录到 Azure** 以登录到 Azure 帐户。 如果还没有 Azure 帐户，请选择 **“创建 Azure 帐户”创建一个帐户**。 按照提供的步骤设置帐户。
5. 登录后，会看到 Azure 存储帐户显示在扩展中。 如果还没有存储帐户，请在命令面板中使用 **“创建存储帐户** ”选项创建一个。 仅使用“a-z”和“0-9”将存储帐户命名为全局唯一名称。 请注意，默认情况下，这将创建一个存储帐户和一个名称相同的资源组。 它会自动将存储帐户放入美国西部。 可以通过 [Azure 帐户](https://portal.azure.com/)进行联机调整。
6. 选择并按住 (右键单击存储帐户) ，然后选择 **“配置静态网站**”。 系统将要求输入索引文档名称和 404 文档名称。 将索引文档名称从默认 `index.html` 值更改为 **`taskpane.html`**。 也可以更改 404 文档名称，但不需要更改。
7. 再次选择并按住 (右键单击) 存储，这次选择 **“浏览静态网站**”。 在打开的浏览器窗口中，复制网站 URL。
8. 在 VS Code 中，打开项目的清单文件 () `manifest.xml` ，并更改对 localhost URL (的任何引用，例如 `https://localhost:3000`) 复制的 URL。 此终结点是新创建的存储帐户的静态网站 URL。 保存对清单文件所做的更改。
9. 打开命令行提示符并导航到加载项项目的根目录。 然后运行以下命令，为生产部署准备所有文件。

    ```command&nbsp;line
    npm run build
    ```

    生成完成后，加载项项目的根目录中的 **dist** 文件夹将包含要在后续步骤中部署的文件。

10. 若要部署，请选择文件资源管理器，选择并按住 (右键单击 **dist** 文件夹) ，然后 **通过 Azure 存储选择“部署到静态网站**”。 出现提示时，选择之前创建的存储帐户。

    :::image type="content" source="../images/deploy-to-static-website.png" alt-text="选择 dist 文件夹、右键单击并选择通过 Azure 存储部署到静态网站。":::

11. 部署完成后，右键单击之前创建的存储帐户，然后选择 **“浏览静态网站**”。 这将打开静态网站并显示任务窗格。

## <a name="deploy-custom-functions-for-excel"></a>为 Excel 部署自定义函数

如果加载项具有自定义函数，则可通过其他几个步骤在 Azure 存储帐户上启用它们。 首先，启用 CORS，以便 Office 可以访问 functions.json 文件。

1. 右键单击 Azure 存储帐户，然后 **选择“在门户中打开**”。
1. 在“设置”组中，选择 **(CORS) 的资源共享**。 还可以使用搜索框来查找此信息。
1. 使用以下设置创建新的 CORS 规则。

    |属性        |值                        |
    |----------------|-----------------------------|
    |允许的源 | \*                          |
    |允许的方法 | GET                         |
    |允许的标头 | \*                          |
    |公开的标头 | Access-Control-Allow-Origin |
    |最大年龄         | 200                         |

1. 选择“**保存**”。

> [!CAUTION]
> 此 CORS 配置假定服务器上的所有文件都公开提供给所有域。  

接下来，为 JSON 文件添加 MIME 类型。

1. 在名为 **web.config** 的 /src 文件夹中创建新文件。
1. 插入以下 XML 并保存文件。

    ```xml
    <?xml version="1.0"?>
    <configuration>
      <system.webServer>
        <staticContent>
          <mimeMap fileExtension=".json" mimeType="application/json" />
        </staticContent>
      </system.webServer>
    </configuration> 
    ```

1. 打开 **webpack.config.js** 文件。
1. 在运行生成时，在列表中 `plugins` 添加以下代码以将web.config复制到捆绑包中。

    ```javascript
    new CopyWebpackPlugin({
      patterns: [
      {
        from: "src/web.config",
        to: "src/web.config",
      },
     ],
    }),
    ```

1. 打开命令行提示符并转到加载项项目的根目录。 然后，运行以下命令准备所有文件以进行部署。

    ```command&nbsp;line
    npm run build
    ```

    生成完成后，外接程序项目根目录中的 **dist** 文件夹将包含要部署的文件。

1. 若要部署，**请在文件资源管理器** 中选择并按住 (或右键单击 **dist** 文件夹) ，然后 **通过 Azure 存储选择“部署到静态网站**”。 出现提示时，选择之前创建的存储帐户。 如果已部署 **dist** 文件夹，如果想要使用最新更改覆盖 Azure 存储中的文件，系统会提示你。

## <a name="see-also"></a>另请参阅

- [使用 Visual Studio Code 开发 Office 加载项](../develop/develop-add-ins-vscode.md)
- [部署和发布 Office 外接程序](../publish/publish.md)
- [跨源资源共享 (CORS) 对 Azure 存储的支持](/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)

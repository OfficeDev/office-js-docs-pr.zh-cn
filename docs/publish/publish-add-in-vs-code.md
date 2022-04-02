---
title: 使用 Azure 和 azure Visual Studio Code加载项
description: 如何使用加载项和加载项Visual Studio Code Azure Active Directory
ms.date: 02/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1559c74493a511bb964fd43159069c1e9e78365e
ms.sourcegitcommit: 8f7d84c33c61c9f724f956740ced01a83f62ddc6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/01/2022
ms.locfileid: "64605519"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>发布使用 Visual Studio Code 开发的加载项

本文介绍如何发布使用 Yeoman 生成器创建并使用 [Visual Studio Code (VS Code)](https://code.visualstudio.com) 或任何其他编辑器开发的 Office 加载项。

> [!NOTE]
> 要了解如何发布使用 Visual Studio 创建的 Office 加载项，请参阅[使用 Visual Studio 发布加载项](package-your-add-in-using-visual-studio.md)。

## <a name="publishing-an-add-in-for-other-users-to-access"></a>发布加载项供其他人用户访问

Office 加载项 包含 Web 应用程序和清单文件。Web 应用程序定义加载项的用户界面和功能，清单指定 Web 应用程序的位置并定义加载项的设置和功能。

开发时，可以在本地 Web 服务器上运行 `localhost` 加载项， () 。 准备好发布它供其他用户访问时，需要部署 Web 应用程序并更新清单以指定已部署应用程序的 URL。

当外接程序根据需要工作时，可以使用 Visual Studio Code 扩展直接Azure 存储它。

## <a name="using-visual-studio-code-to-publish"></a>使用Visual Studio Code发布

>[!NOTE]
> 这些步骤仅适用于使用 Yeoman 生成器创建的项目。

1. 从项目根文件夹中打开项目，Visual Studio Code (VS Code) 。
2. 从"扩展"视图中VS Code，搜索并Azure 存储扩展。
3. 安装后，Azure 图标将添加到活动栏。 选择它以访问扩展。 如果活动栏处于隐藏状态，你将无法访问扩展。 通过选择"显示活动栏"> **">"活动栏"来显示活动栏**。
4. 在扩展中时，通过选择"登录到 Azure" **登录到 Azure 帐户**。 如果还没有 Azure 帐户，也可以选择"创建免费的 Azure 帐户"，创建 **Azure 帐户**。 按照提供的步骤设置帐户。
5. 登录 Azure 帐户后，你将看到 Azure 存储帐户显示在扩展中。 如果还没有存储帐户，则使用命令调色板中的"创建存储 **帐户**"选项创建一个存储帐户。 将存储帐户命名为全局唯一名称，仅使用"a-z"和"0-9"。 请注意，默认情况下，这将创建一个存储帐户和一个同名的资源组。 它会自动将存储帐户置于美国西部。 这可以通过 Azure 帐户在线 [调整](https://portal.azure.com/)。
6. 选择并按住 (右键) 存储帐户，选择" **配置静态网站"**。 将要求您输入索引文档名称和 404 文档名称。 将索引文档名称从默认更改为 `index.html` **`taskpane.html`**。 您还可以更改 404 文档名称，但不要求更改。
7. 选择并按住 (再次右键) 存储"，这次选择"浏览 **静态网站"**。 从打开的浏览器窗口中复制网站 URL。
8. In VS Code， open your project's manifest file (`manifest.xml`) and change any reference to your localhost URL (such as `https://localhost:3000`) to the URL you've copied. 此终结点是新创建的存储帐户的静态网站 URL。 保存对清单文件所做的更改。
9. 打开命令行提示符并导航到加载项项目的根目录。 然后运行以下命令以准备用于生产部署的所有文件。

    ```command&nbsp;line
    npm run build
    ```

    生成完成后，加载项项目的根目录中的 **dist** 文件夹将包含要在后续步骤中部署的文件。

10. 若要部署，请选择"文件资源管理器"，选择并按住 (右键单击") **"，** 然后选择"通过"Azure 存储 **部署到静态网站"**。 当系统提示时，选择之前创建的存储帐户。

    ![部署到静态网站。](../images/deploy-to-static-website.png)

11. 部署完成后 **，将显示"** 浏览到网站"消息，您可以选择该消息打开已部署应用代码的主终结点。

## <a name="deploy-custom-functions-for-excel"></a>部署自定义函数Excel

如果您的外接程序具有自定义函数，则还有一些附加步骤可将其启用Azure 存储帐户。 首先，你需要启用 CORS，以便Office functions.json 文件。

1. 右键单击 Azure 存储帐户，然后选择" **在门户中打开"**。
1. 在"设置组中，选择"资源 **共享 (CORS) "**。 您还可以使用搜索框查找此内容。
1. 创建具有以下设置的新 CORS 规则。

    |属性        |值                        |
    |----------------|-----------------------------|
    |允许的来源 | \*                          |
    |允许的方法 | GET                         |
    |允许的标头 | \*                          |
    |公开的标头 | Access-Control-Allow-Origin |
    |最长年龄         | 200                         |

1. 选择“**保存**”。

> [!CAUTION]
> 此 CORS 配置假定您的服务器上的所有文件都公开可用于所有域。  

接下来，你需要为 JSON 文件添加 MIME 类型。

1. 在名为web.config的 /src **文件夹中web.config**。
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
1. 在 列表中添加以下代码 `plugins` ，以在生成运行时将 web.config复制到捆绑包中。

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

1. 打开命令行提示符并转到加载项项目的根目录。 然后运行以下命令以准备所有文件以部署。

    ```command&nbsp;line
    npm run build
    ```

    生成完成后，外接程序项目的根目录中的 **dist** 文件夹将包含您将部署的文件。

1. 若要部署，请选择"文件资源管理器"，选择并按住 (右键单击") **"，** 然后选择"通过"Azure 存储 **部署到静态网站"**。 当系统提示时，选择之前创建的存储帐户。 如果你已部署 **dist** 文件夹，当你想要用最新更改覆盖 Azure 存储中的文件时，系统将提示你。

## <a name="see-also"></a>另请参阅

- [使用 Visual Studio Code 开发 Office 加载项](../develop/develop-add-ins-vscode.md)
- [部署和发布 Office 外接程序](../publish/publish.md)
- [跨源资源共享 (CORS) 支持Azure 存储](/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)

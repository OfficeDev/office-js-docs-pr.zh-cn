---
title: 使用 Visual Studio Code 和 Azure 发布加载项
description: 如何使用 Visual Studio Code 和 Azure Active Directory 发布加载项
ms.date: 09/07/2022
ms.custom: vscode-azure-extension-update-completed
ms.localizationpriority: medium
ms.openlocfilehash: f1e931042f7796333bbc75b2666f3564e63b9f7b
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467189"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>发布使用 Visual Studio Code 开发的加载项

本文介绍如何发布使用 Yeoman 生成器创建并使用 [Visual Studio Code (VS Code)](https://code.visualstudio.com) 或任何其他编辑器开发的 Office 加载项。

> [!NOTE]
> 要了解如何发布使用 Visual Studio 创建的 Office 加载项，请参阅[使用 Visual Studio 发布加载项](package-your-add-in-using-visual-studio.md)。

## <a name="publishing-an-add-in-for-other-users-to-access"></a>发布加载项供其他人用户访问

An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.

开发时，可以在本地 Web 服务器上运行加载项 () `localhost` 。 准备好将其发布供其他用户访问时，需要部署 Web 应用程序并更新清单以指定已部署应用程序的 URL。

当加载项按需工作时，可以使用 Azure 存储扩展直接通过Visual Studio Code发布。

## <a name="using-visual-studio-code-to-publish"></a>使用 Visual Studio Code 发布

>[!NOTE]
> 这些步骤仅适用于使用 Yeoman 生成器创建的项目。

1. 在 VISUAL STUDIO CODE (VS Code) 中从其根文件夹打开项目。
1. 选择“ **查看>扩展 (** Ctrl+Shift+X) 以打开”扩展“视图。
1. 搜索 **Azure 存储** 扩展并安装它。
1. 安装后，Azure 图标将添加到 **活动栏**。 选择它以访问扩展。 如果 **活动栏** 处于隐藏状态，请选择 **“视图”>“外观”>活动栏** 将其打开。
1. 选择 **登录到 Azure** 以登录到 Azure 帐户。 如果还没有 Azure 帐户，请选择 **“创建 Azure 帐户”创建一个帐户**。 按照提供的步骤设置帐户。

    :::image type="content" source="../images/azure-extension-sign-in.png" alt-text="登录到 Azure 扩展中选择的 Azure 按钮。":::

1. 登录后，会看到 Azure 存储帐户显示在扩展中。 如果还没有存储帐户，请在命令面板中使用 **“创建存储帐户** ”选项创建一个。 仅使用“a-z”和“0-9”将存储帐户命名为全局唯一名称。 请注意，默认情况下，这将创建一个存储帐户和一个名称相同的资源组。 它会自动将存储帐户放入美国西部。 可以通过 [Azure 帐户](https://portal.azure.com/)进行联机调整。

    :::image type="content" source="../images/azure-extension-create-storage-account.png" alt-text="选择存储帐户>在 Azure 扩展中创建存储帐户。":::

1. 右键单击存储帐户，然后选择 **“配置静态网站**”。 系统将要求输入索引文档名称和 404 文档名称。 将索引文档名称从默认 `index.html` 值更改为 **`taskpane.html`**。 也可以更改 404 文档名称，但不需要更改。
1. 再次右键单击存储帐户，这次选择 **“浏览静态网站**”。 在打开的浏览器窗口中，复制网站 URL。
1. 打开项目的清单文件 () `manifest.xml` ，并更改对 localhost URL (的所有引用，例如 `https://localhost:3000`) 复制的 URL。 此终结点是新创建的存储帐户的静态网站 URL。 保存对清单文件所做的更改。
1. 打开命令行提示符或终端窗口，然后转到加载项项目的根目录。 运行以下命令，为生产部署准备所有文件。

    ```command&nbsp;line
    npm run build
    ```

    生成完成后，加载项项目的根目录中的 **dist** 文件夹将包含要在后续步骤中部署的文件。

1. 在 VS Code 中，转到资源管理器并右键单击 **dist** 文件夹，然后 **通过 Azure 存储选择“部署到静态网站**”。 出现提示时，选择之前创建的存储帐户。

    :::image type="content" source="../images/deploy-to-static-website.png" alt-text="选择 dist 文件夹，右键单击，然后通过 Azure 存储选择“部署到静态网站”。":::

1. 部署完成后，右键单击之前创建的存储帐户，然后选择 **“浏览静态网站**”。 这将打开静态网站并显示任务窗格。

1. 最后，从刚部署的静态网站 [旁加载清单文件](../testing/sideload-office-add-ins-for-testing.md) 和加载项。

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

1. 若要部署，请在 VS Code **Explorer** 中右键单击 **dist** 文件夹，然后 **通过 Azure 存储选择“部署到静态网站**”。 出现提示时，选择之前创建的存储帐户。 如果已部署 **dist** 文件夹，如果想要使用最新更改覆盖 Azure 存储中的文件，系统会提示你。

## <a name="see-also"></a>另请参阅

- [使用 Visual Studio Code 开发 Office 加载项](../develop/develop-add-ins-vscode.md)
- [部署和发布 Office 外接程序](../publish/publish.md)
- [跨源资源共享 (CORS) 对 Azure 存储的支持](/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)

---
title: 在 Office 网页版中旁加载 Office 加载项进行测试
description: 通过旁Office在Office web 版外接程序中测试外接程序。
ms.date: 04/14/2021
localization_priority: Normal
ms.openlocfilehash: e830ccbb6a4e325d6d70c3612492009b5e3d1570
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077216"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a>在 Office 网页版中旁加载 Office 加载项进行测试

旁加载加载项时，无需先将加载项放在加载项目录中，即可安装加载项。 在测试和开发外接程序时，这非常有用，因为你可以看到外接程序的显示和运行方式。

旁加载外接程序时，外接程序的清单存储在浏览器的本地存储中，因此，如果您清除浏览器的缓存或切换到其他浏览器，您必须再次旁加载外接程序。

旁加载因主机应用程序而异 (例如，Excel) 。

> [!NOTE]
> 如本文所述，在 Excel、OneNote、PowerPoint 和 Word 上支持旁加载。 若要旁加载 Outlook 外接程序，请参阅[旁加载 Outlook 外接程序进行测试](../outlook/sideload-outlook-add-ins-for-testing.md)。

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a>在 Office 网页版中旁加载 Office 加载项

此过程仅支持 **Excel、OneNote、PowerPoint** 和 **Word。**  有关其他主机应用程序，请参阅以下部分中的手动旁加载说明。 此示例项目假定你正在使用使用[Yeoman](https://github.com/OfficeDev/generator-office)生成器为加载项创建Office项目。

1. 打开[Office web 版](https://office.live.com/)。 使用"**创建"** 选项，在"新建"Excel、OneNote、PowerPoint或 Word **中****创建文档**。 在此新文档中，选择功能 **区** 中的"共享"，选择" **复制链接**"，然后复制 URL。

2. 在 yo office 项目文件的根目录中，打开package.js **on** 文件。 在此 **文件的"配置** "部分，创建 `"document"` 一个属性。 粘贴您复制的 URL 作为属性的值 `"document"` 。 例如，你的将如下所示：

    ```json
      "config": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > 如果创建的加载项不是使用 Yeoman 生成器，可以通过将以下内容附加到现有 URL，将查询参数添加到文档的 URL：

    - 开发服务器端口，例如 `&wdaddindevserverport=3000` 。
    - 清单文件名，例如 `&wdaddinmanifestfile=manifest1.xml` 。
    - 清单 GUID，例如 `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143` 。

    > 如果你使用的是 Yeoman 生成器，则不需要添加此信息，因为 Yeoman 工具会自动附加此信息。
    > 请注意，在这两种情况下，只能从 localhost 加载清单。

3. 在从项目的根目录开始的命令行中，运行以下命令： `npm run start:web` 。

4. 首次使用此方法在 Web 上旁加载外接程序时，你将看到一个对话框，要求您启用开发人员模式。 选中"现在启用 **开发人员模式"复选框，** 然后选择"确定 **"。**

5. 你将看到第二个对话框，询问您是否希望从Office注册外接程序清单。 应选择"**是"。**

6. 已安装您的外接程序。 如果是加载项命令，它应显示在功能区或上下文菜单上。 如果是任务窗格加载项，应显示任务窗格。

## <a name="sideload-an-office-add-in-in-office-on-the-web-manually"></a>手动Office加载项旁Office web 版加载项

此方法不使用命令行，只能在主机应用程序（如 Excel) ）内使用命令 (实现。

1. 打开[Office web 版](https://office.live.com/)。 在"文档"Excel、Word或 PowerPoint **中打开文档**。  在"**外接程序**"部分的功能区上的"插入"选项卡上，Office **外接程序"。**

1. 在 **"Office** 外接程序"对话框中，选择"**我的** 外接程序"选项卡，选择"管理 **我的** 外接程序"，Upload"**我的外接程序"。**

    ![the Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in".](../images/office-add-ins-my-account.png)

1. **转到** 加载项清单文件，再选择“上传”。

    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

1. 验证是否已安装外接程序。例如，如果它是一个外接程序命令，它应显示在功能区或上下文菜单上。如果它是一个任务窗格外接程序，则应显示窗格。

> [!NOTE]
> 若要使用Office EdgeHTML Microsoft Edge WebView (测试加载项) ，需要执行其他配置步骤。 在命令Windows中，运行以下行： `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes` 。 使用基于 Office 的边缘 WebView2 时Chromium此要求。 有关详细信息，请参阅Office[使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

## <a name="sideload-an-office-add-in"></a>旁加载Office加载项

1. 登录到你的 Microsoft 365 帐户。

2. 打开工具栏Launcher"应用"菜单，选择"Excel、Word 或 **PowerPoint"，** 然后创建新文档。

3. 步骤 3 - 6 与上一部分 **在 Office 网页版中旁加载 Office 加载项** 相同。

## <a name="sideload-an-add-in-when-using-visual-studio"></a>使用 Visual Studio 时旁加载加载项

如果你使用 Visual Studio 开发外接程序，旁加载的过程类似于手动旁加载到 Web。 唯一的区别是，必须更新清单中 **SourceURL** 元素的值以包含部署加载项位置的完整 URL。

> [!NOTE]
> 虽然可以将加载项从 Visual Studio 旁加载到 Office 网页版，但无法从 Visual Studio 调试它们。 若要进行调试，需要使用浏览器调试工具。 有关详细信息，请参阅[在 Office 网页版中调试加载项](debug-add-ins-in-office-online.md)。

1. 在 Visual Studio 中，通过选择 **视图** > **属性窗口** 来显示 **属性** 窗口。
2. 在 **解决方案资源管理器** 中，选择 Web 项目。 这将在 **属性** 窗口中显示项目的属性。
3. 在“属性”窗口中复制 **SSL URL**。
4. 在加载项项目中，打开清单 XML 文件。 请确保正在编辑源 XML。 对于某些项目类型，Visual Studio 将打开 XML 的可视视图，它不适用于下一步骤。
5. 使用刚复制的 SSL URL 来搜索和替换 **~remoteAppUrl/** 的所有实例。 将看到多个替换，具体取决于项目类型。将显示新 URL，类似于 `https://localhost:44300/Home.html`。
6. 保存 XML 文件。
7. 右键单击 Web 项目，然后选择 **调试** > **启动新实例**。 这将在不启动 Office 的情况下运行 Web 项目。
8. 从 Office 网页版，使用之前[在 Office 网页版中加载 Office 加载项](#sideload-an-office-add-in-in-office-on-the-web)中所述的步骤旁加载加载项。

## <a name="remove-a-sideloaded-add-in"></a>删除旁加载的外接程序

可以通过清除浏览器的缓存来删除以前旁加载的外接程序。 如果您更改外接程序的清单 (例如，更新图标的文件名或外接程序命令) 的文本，您可能需要清除[Office](clear-cache.md)缓存，然后使用更新后的清单重新旁加载外接程序。 执行此操作后，Office 将按照更新清单中所述的方式呈现该加载项。

## <a name="see-also"></a>另请参阅

- [在 iPad 和 Mac 上旁加载 Office 加载项](sideload-an-office-add-in-on-ipad-and-mac.md)
- [旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)
- [清除 Office 缓存](clear-cache.md)

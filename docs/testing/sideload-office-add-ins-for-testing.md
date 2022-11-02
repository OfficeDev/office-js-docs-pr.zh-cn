---
title: 将 Office 加载项旁加载到Office web 版
description: 通过旁加载在 Office web 版 中测试 Office 加载项。
ms.date: 09/02/2022
ms.localizationpriority: medium
ms.openlocfilehash: 128e3537ac0ece5b7574dfec6d9d5c67b8d95a7b
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810379"
---
# <a name="sideload-office-add-ins-to-office-on-the-web"></a>将 Office 加载项旁加载到Office web 版

旁加载加载项时，无需先将其放入外接程序目录中即可安装加载项。 这在测试和开发加载项时很有用，因为可以看到加载项的显示方式和功能。

在 Web 上旁加载加载项时，加载项清单存储在浏览器的本地存储中，因此，如果清除浏览器的缓存或切换到其他浏览器，则必须再次旁加载加载项。

在 Web 上旁加载加载项的步骤因以下因素而异。

- 主机应用程序 (，例如 Excel、Word、Outlook) 
- 创建外接程序项目的工具 (例如 Visual Studio 和 Office 外接程序的 Yeoman 生成器，或者两者都没有) 
- 是使用 Microsoft 帐户还是 Microsoft 365 租户中的帐户旁加载到Office web 版

在以下列表中，转到与方案匹配的部分或文章。 请注意，列表中的第一个方案适用于 Outlook 加载项。其余方案适用于非 Outlook 加载项。

- 如果要旁加载 Outlook 加载项，请参阅 [旁加载 Outlook 加载项以进行测试](../outlook/sideload-outlook-add-ins-for-testing.md)一文。
- 如果使用 Office 加载项的 [Yeoman 生成器创建了加载项](../develop/yeoman-generator-overview.md)，请参阅[将 Yeoman 创建的加载项旁加载到Office web 版](#sideload-a-yeoman-created-add-in-to-office-on-the-web)。
- 如果使用 Visual Studio 创建了加载项，请参阅 [使用 Visual Studio 时在 Web 上旁加载加载项](#sideload-an-add-in-on-the-web-when-using-visual-studio)。
- 对于所有其他情况，请参阅以下部分之一。

  - 如果要使用 Microsoft 帐户旁加载到Office web 版，请参阅[手动旁加载加载项到Office web 版](#manually-sideload-an-add-in-to-office-on-the-web)。
  - 如果要使用 Microsoft 365 租户中的帐户旁加载到Office web 版，请参阅[将加载项旁加载到 Microsoft 365](#sideload-an-add-in-to-microsoft-365)。

## <a name="sideload-a-yeoman-created-add-in-to-office-on-the-web"></a>将 Yeoman 创建的加载项旁加载到 Office web 版

此过程仅支持 **Excel**、 **OneNote**、 **PowerPoint** 和 **Word** 。 此示例项目假定你使用的是 [使用 Office 外接程序的 Yeoman 生成器](../develop/yeoman-generator-overview.md)创建的项目。

1. 打开 [Office web 版](https://office.live.com/) 或 OneDrive。 使用 **“创建”** 选项，在 **Excel**、 **OneNote**、 **PowerPoint** 或 **Word** 中创建文档。 在此新文档中，选择“ **共享**”，选择“ **复制链接**”，然后复制 URL。

1. 在从项目的根目录开始的命令行中，运行以下命令。 将“{url}”替换为复制的 URL。

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. 首次使用此方法旁加载 Web 上的加载项时，会看到一个对话框，要求启用开发人员模式。 选中“ **立即启用开发人员模式** ”复选框，然后选择 **“确定**”。

1. 你将看到另一个对话框，询问是否要从计算机注册 Office 外接程序清单。 选择“是”。

1. 加载项已安装。 如果它具有外接程序命令，则它应显示在功能区或上下文菜单上。 如果它是任务窗格加载项，而没有任何加载项命令，则任务窗格应显示。

## <a name="sideload-an-add-in-on-the-web-when-using-visual-studio"></a>使用 Visual Studio 时在 Web 上旁加载加载项

如果使用 Visual Studio 开发加载项，请按 **F5** 在 *桌面* Office 中打开 Office 文档，创建空白文档，然后旁加载加载项。 如果要旁加载到 *Office web 版*，旁加载的过程类似于手动旁加载到 Web。 唯一的区别是，必须更新清单中 **SourceURL** 元素的值，并可能更新其他元素的值，以包括部署外接程序的完整 URL。

1. 在 Visual Studio 中，选择“ **查看** > **属性窗口**”。

1. 在 **解决方案资源管理器** 中，选择 Web 项目。 这会在“属性”窗口中显示项目 **的属性** 。

1. 在“属性”窗口中复制 **SSL URL**。

1. 在加载项项目中，打开清单 XML 文件。 请确保正在编辑源 XML。 对于某些项目类型，Visual Studio 将打开 XML 的可视视图，该视图在下一步中不起作用。

1. 使用刚复制的 SSL URL 来搜索和替换 **~remoteAppUrl/** 的所有实例。 你将看到多个替换项，具体取决于项目类型，并且新 URL 将类似于 `https://localhost:44300/Home.html`。

1. **保存** XML 文件。

1. 在 **გადაწყვეტების მნახველი** 中，打开 Web 项目的上下文菜单 (例如，右键单击它) 然后选择 **“调试** > **启动新实例**”。 这会在不启动 Office 的情况下运行 Web 项目。

1. 从Office web 版，使用手动旁加载加载项[到Office web 版中所述的步骤旁加载加载项](#manually-sideload-an-add-in-to-office-on-the-web)。

## <a name="manually-sideload-an-add-in-to-office-on-the-web"></a>将加载项手动旁加载到 Office web 版

此方法不使用命令行，只能在主机应用程序 (（如 Excel) ）中使用命令来完成。

1. 打开[Office web 版](https://office.com/)。 在 **Excel**、 **OneNote**、 **PowerPoint** 或  **Word** 中打开文档。 

1. 在“ **插入** ”选项卡上的“ **加载项** ”部分中，选择“ **Office 加载项**”。

1. 在 **“Office 加载项** ”对话框中，选择“ **我的外接程序** ”选项卡，选择“ **管理我的外接程序”**，然后选择 **“上传我的外接程序**”。

    ![Office 加载项对话框的右上方有一个下拉列表，上面写着“管理我的加载项”，下方有一个带有“上传我的外接程序”选项的下拉列表。](../images/office-add-ins-my-account.png)

1. **转到** 加载项清单文件，再选择“上传”。

    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

1. 验证是否已安装外接程序。 例如，如果它具有外接程序命令，则它应显示在功能区或上下文菜单上。 如果它是没有外接程序命令的任务窗格加载项，则应显示任务窗格。

> [!NOTE]
> 若要使用原始 WebView (EdgeHTML) 通过 Microsoft Edge 测试 Office 加载项，需要执行其他配置步骤。 在 Windows 命令提示符中，运行以下行： `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`。 当 Office 使用基于 Chromium 的 Edge WebView2 时，这不是必需的。 有关详细信息，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

[!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

## <a name="sideload-an-add-in-to-microsoft-365"></a>将加载项旁加载到 Microsoft 365

1. 登录到 Microsoft 365 帐户。

1. 打开工具栏左端的应用程序启动器，选择 **“Excel**”、“ **OneNote**”、“ **PowerPoint”** 或 **“Word**”，然后创建新文档。

1. 在“ **插入** ”选项卡上，选择“ **加载项** ”按钮。

1. 按照[手动旁加载加载项到Office web 版](#manually-sideload-an-add-in-to-office-on-the-web)部分的步骤 3 - 5 进行操作。

## <a name="remove-a-sideloaded-add-in"></a>删除旁加载加载项

若要删除旁加载到Office web 版的加载项，只需清除浏览器的缓存即可。 例如，如果对加载项的清单进行了更改 (更新图标的文件名或外接程序命令的文本) ，则可能需要清除浏览器的缓存，然后使用更新的清单重新旁加载加载项。 这样做允许Office web 版呈现加载项，如更新的清单所述。

## <a name="see-also"></a>另请参阅

- [在 Mac 上旁加载 Office 加载项](sideload-an-office-add-in-on-mac.md)
- [在 iPad 上旁加载 Office 加载项](sideload-an-office-add-in-on-ipad.md)
- [旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)
- [清除 Office 缓存](clear-cache.md)

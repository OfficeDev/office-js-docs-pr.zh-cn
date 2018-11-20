---
title: 在 Office Online 中旁加载 Office 加载项以供测试
description: 通过旁加载在 Office Online 中测试 Office 加载项
ms.date: 10/19/2018
ms.openlocfilehash: 94138cd0a22f053a9471bf905b8d0838dead15cf
ms.sourcegitcommit: 3a808cf39cbc77056968d53a5957462371ad83a1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2018
ms.locfileid: "25911226"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a>在 Office Online 中旁加载 Office 加载项以供测试

可以通过使用旁加载来安装 Office 加载项以供测试，而无需先将它添加到加载项目录中。 在 Office 365 或 Office Online 中都可以进行旁加载。 该过程使用的两个平台略有不同。 

当旁加载外接程序时，外接程序清单存储在浏览器的本地存储区中，因此如果清除浏览器的缓存，或切换到另一个浏览器，就必须再次旁加载该外接程序。


> [!NOTE]
> 如本文所述，Word、Excel 和 PowerPoint 支持旁加载。若要旁加载 Outlook 外接程序，请参阅[旁加载 Outlook 外接程序进行测试](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)。

下面的视频逐步展示了如何在 Office 桌面或 Office Online 上旁加载加载项。  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-365"></a>在 Office 365 上旁加载 Office 加载项


1. 登录 Office 365 帐户。
    
2. 打开工具栏最左端的应用启动器，选择“Excel”****、“Word”**** 或“PowerPoint”****，再新建文档。
    
3. 打开功能区上的“**插入**”选项卡，然后在“**外接程序**”部分中，选择“**Office 外接程序**”。
    
4. 在“Office 加载项”**** 对话框中，依次选择“我的组织”**** 选项卡和“上传我的加载项”****。
    
    ![标题为“Office 加载项”的对话框，左上角附近有链接“上传我的加载项”](../images/office-add-ins.png)

5.  **转到**加载项清单文件，再选择“上传”****。
    
    ![包含“浏览”、“上传”和“取消”按钮的“上传加载项”对话框](../images/upload-add-in.png)

6. 验证是否已安装外接程序。例如，如果它是一个外接程序命令，它应显示在功能区或上下文菜单上。如果它是一个任务窗格外接程序，则应显示窗格。
    

## <a name="sideload-an-office-add-in-in-office-online"></a>在 Office Online 中旁加载 Office 加载项


1. 打开 [Microsoft Office Online](https://office.live.com/)。
    
2. 在“**立即开始使用在线应用**”中，选择 **Excel**、**Word** 或 **PowerPoint**；然后打开一个新文档。
    
3. 打开功能区上的“**插入**”选项卡，然后在“**外接程序**”部分中，选择“**Office 外接程序**”。
    
4. 在“Office 加载项”**** 对话框中，依次选择“我的加载项”**** 选项卡、“管理我的加载项”**** 和“上传我的加载项”****。
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

5.  **转到**加载项清单文件，再选择“上传”****。
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

6. 验证是否已安装外接程序。例如，如果它是一个外接程序命令，它应显示在功能区或上下文菜单上。如果它是一个任务窗格外接程序，则应显示窗格。

> [!NOTE]
>若要使用 Edge 测试 Office 加载项，请在 Edge 搜索栏中输入“** about：flags **”以调出“开发人员设置”选项。  选中“**允许本地主机环回**”选项，然后重新启动 Edge。

>    ![选中此框后，Edge 会允许本地主机环回选项。](../images/allow-localhost-loopback.png)

## <a name="sideload-an-add-in-when-using-visual-studio"></a>使用 Visual Studio 时旁加载加载项

如果使用 Visual Studio 来开发加载项，则旁加载的过程类似。 唯一的区别是，必须更新清单中 **SourceURL** 元素的值以包含部署加载项位置的完整 URL。

> [!NOTE]
> 虽然可以将加载项从 Visual Studio 旁加载到 Office Online，但无法从 Visual Studio 调试它们。 若要进行调试，需要使用浏览器调试工具。 有关详细信息，请参阅[在 Office Online 中调试加载项](debug-add-ins-in-office-online.md)。

1. 在 Visual Studio 中，通过选择**视图** -> **属性窗口**来显示**属性**窗口。
2. 在**解决方案资源管理器**中，选择 Web 项目。 这将在**属性**窗口中显示项目的属性。
3. 在“属性”窗口中复制 **SSL URL**。
4. 在加载项项目中，打开清单 XML 文件。 请确保正在编辑源 XML。 对于某些项目类型，Visual Studio 将打开 XML 的可视视图，它不适用于下一步骤。
5. 使用刚复制的 SSL URL 来搜索和替换 **~remoteAppUrl/** 的所有实例。 将看到多个替换，具体取决于项目类型。将显示新 URL，类似于 `https://localhost:44300/Home.html`。
6. 保存 XML 文件。
7. 右键单击 Web 项目，然后选择**调试** -> **启动新实例**。 这将在不启动 Office 的情况下运行 Web 项目。
8. 从 Office Online，使用之前[在 Office Online 中加载 Office 加载项](#sideload-an-office-add-in-in-office-online)中所述的步骤旁加载加载项。

---
title: 旁加载 Office 加载项以供测试
description: ''
ms.date: 10/17/2018
localization_priority: Priority
ms.openlocfilehash: 7d7aa136e7268183c5a064f16315d433ee53a38b
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388841"
---
# <a name="sideload-office-add-ins-for-testing"></a>旁加载 Office 加载项以供测试

你可以安装 Office 外接程序以在 Windows 上运行的 Office 客户端中进行测试（通过使用共享文件夹，以将清单发布到网络文件共享）。

> [!NOTE]
> 如果你的加载项项目是使用 [**yo office** 工具](https://github.com/OfficeDev/generator-office)创建的，则可以使用另一种方法来旁加载它，这可能对你有用。 有关详细信息，请参阅[使用旁加载命令旁加载 Office 加载项](sideload-office-addin-using-sideload-command.md)。

本文仅适用于在 Windows 上测试 Word、Excel 或 PowerPoint 加载项。 如果要在其他平台上进行测试或要测试 Outlook 加载项，请参阅以下主题之一以旁加载你的加载项：

- [在 Office Online 中旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
- [在 iPad 和 Mac 上旁加载 Office 外接程序进行测试](sideload-an-office-add-in-on-ipad-and-mac.md)
- [旁加载 Outlook 加载项以供测试](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)


下面的视频逐步展示了如何使用共享文件夹目录在 Office 桌面或 Office Online 上旁加载加载项。  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a>共享文件夹

1. 在想要托管外接程序的 Windows 计算机上，转到你想用作共享文件夹目录的文件夹的父文件夹或驱动器号。

2. 打开要用作共享文件夹目录的文件夹的上下文菜单（右键单击该文件夹），然后选择“**属性**”。

3. 在“**属性**”对话框窗口中，打开“**共享**”选项卡，然后选择“**共享**”按钮。

    ![已突出显示“共享”选项卡和“共享”按钮的文件夹“属性”对话框](../images/sideload-windows-properties-dialog.png)

4. 在**网络访问**对话框窗口中，添加你自己以及要与其共享加载项的任何其他用户和/或组。 你至少需要对该文件夹的**读/写**权限。 选择要与其共享的人员后，请选择“**共享**”按钮。

5. 当你看到确认**你的文件夹已共享**的消息时，请记下紧跟文件夹名称显示的完整网络路径。 （当你[将共享文件夹指定为受信任的目录](#specify-the-shared-folder-as-a-trusted-catalog)时，你需要将此值输入为**目录UR **，如本文下一节所述。）选择“**完成**”按钮以关闭“**网络访问**”对话框窗口。

   ![已突出显示共享路径的“网络访问”对话框](../images/sideload-windows-network-access-dialog.png)

6. 选择“**关闭**”按钮以关闭“**属性**”对话框窗口。

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>将共享文件夹指定为受信任的目录
      
1. 在 Excel、Word 或 PowerPoint 中打开一个新文档。
    
2. 选择“**文件**”选项卡，然后选择“**选项**”。
    
3. 选择“**信任中心**”，然后选择“**信任中心设置**”按钮。
    
4. 选择“**受信任的加载项目录**”。
    
5. 在“**目录 Url**”框中，输入你之前[共享](#share-a-folder)的文件夹的完整网络路径。 如果在共享文件夹时未能记下文件夹的完整网络路径，则可以从文件夹的“**属性**”对话框窗口中获取它，如以下屏幕截图所示。 

    ![已突出显示“共享”选项卡和网络路径的文件夹“属性”对话框](../images/sideload-windows-properties-dialog-2.png)
    
6. 在“**目录 Url**”框中输入文件夹的完整网络路径后，选择“**添加目录**”按钮。

7. 选中新添加项目的“**在菜单中显示**”复选框，然后选择“**确定**”按钮以关闭“**信任中心**”对话框窗口。 

    ![已选择目录的“信任中心”对话框](../images/sideload-windows-trust-center-dialog.png)

8. 选择“**确定**”按钮以关闭“**Word 选项**”对话框窗口。

9. 关闭并重新打开 Office 应用程序，以使更改生效。
    

## <a name="sideload-your-add-in"></a>旁加载加载项


1. 放入在共享文件夹目录中进行测试的所有加载项的清单 XML 文件。 请务必将 Web 应用程序本身部署到 Web 服务器。 务必在清单文件的 **SourceLocation** 元素中指定 URL。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. 在 Excel、Word 或 PowerPoint 中，选择功能区上“**插入**”选项卡中的“**我的加载项**”。

3. 在“**Office 外接程序**”对话框的顶部，选择“**共享文件夹**”。

4. 依次选择加载项名称和“确定”****，以插入加载项。


## <a name="see-also"></a>另请参阅

- [验证并排查清单问题](troubleshoot-manifest.md)
- [发布 Office 外接程序](../publish/publish.md)
    

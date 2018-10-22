---
title: 旁加载 Office 外接程序以供测试
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 6ee8e4e9a2413b34cb8991b09d61e16888a0e6a6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640020"
---
# <a name="sideload-office-add-ins-for-testing"></a>旁加载 Office 外接程序以供测试

可以通过将清单发布到网络文件共享来安装 Office 外接程序，以便在 Windows 上运行的 Office 客户端中进行测试（说明如下）。

> [!NOTE]
> 如果外接程序项目是使用 [**yo office** 工具](https://github.com/OfficeDev/generator-office)创建的，那么有一种替代方法可以提供旁加载功能。有关详情，请参阅 [使用 sideload 命令旁加载 Office 外接程序](sideload-office-addin-using-sideload-command.md)。

本文仅适用于在 Windows 上测试 Word、Excel 或 PowerPoint 外接程序。如果要在其他平台上进行测试或想要测试 Outlook 外接程序，请参阅以下主题之一来旁加载外接程序：

- [在 Office Online 中旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
- [在 iPad 和 Mac 上旁加载 Office 外接程序进行测试](sideload-an-office-add-in-on-ipad-and-mac.md)
- [旁加载 Outlook 外接程序以供测试](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)


下面的视频逐步展示了如何在 Office 桌面或 Office Online 上使用共享文件夹目录旁加载外接程序。  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a>共享文件夹

1. 在想要托管外接程序的 Windows 计算机的资源管理器中，转到你想用作共享文件夹目录的文件夹的父文件夹或驱动器号。

2. 打开文件夹的上下文菜单（在文件夹上右键单击）并选择“属性****”。

3. 在“属性”**** 对话框窗口中打开“共享”**** 选项卡，然后选择“共享”**** 按钮。

    ![文件夹“属性”对话框，“共享”选项卡和“共享”按钮被突出显示](../images/sideload-windows-properties-dialog.png)

4. 在“网络访问”**** 对话窗口中，添加自己以及任何想要与其共享外界程序的其他用户和/或组。 你将至少需要该文件夹的**读/写**权限。 选择好要与之共享的人员后，选择“共享”**** 按钮。

5. 当你看到****“你的文件夹已共享”的确认时，记下显示紧接在文件夹名称后的完整网络路径。 （**** 当你“将共享文件夹指定为受信任的目录”时，[](#specify-the-shared-folder-as-a-trusted-catalog)将会需要输入此值作为“目录 Url”，如本文的下一节中所述。)选择“完成”**** 按钮以关闭“网络访问”**** 对话框窗口。

   ![共享路径突出显示的网络访问对话框](../images/sideload-windows-network-access-dialog.png)

6. 选择“关闭”按钮以关闭“属性”对话框窗口。** ** ** **

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>将共享文件夹指定为受信任的目录
      
1. 在 Excel、Word 或 PowerPoint 中打开一个新文档。
    
2. 选择“文件”**** 选项卡，然后选择“选项”****。
    
3. 选择“信任中心”****，然后选择“信任中心设置”**** 按钮。
    
4. 选择“受信任的外接程序目录****”。
    
5. 在“目录 URL****”框中，输入你之前[共享](#share-a-folder) 的文件夹的完整网络路径。 如果你在共享文件夹时没有记下文件夹的完整的网络路径，可以从文件夹的“属性”**** 对话框窗口获取，如以下屏幕截图中所示。 

    ![文件夹“属性”对话框，“共享”选项卡和网络路径被突出显示](../images/sideload-windows-properties-dialog-2.png)
    
6. 在将文件夹的完整网络路径输入到“目录 Url”**** 框中之后，选择“添加目录”**** 按钮。

7. 对于新添加的项目，选择“在菜单中显示”**** 复选框，然后选择“确定” **** 按钮以关闭“信任中心”**** 对话框窗口。 

    ![已选择了目录的“信任中心”对话框](../images/sideload-windows-trust-center-dialog.png)

8. 选择“确定”**按钮**以关闭“Word  选项”**** 对话框窗口。

9. 关闭并重新打开 Office 应用程序，这样你的更改将生效。
    

## <a name="sideload-your-add-in"></a>旁加载外接程序


1. 将进行测试的任意外接程序清单文件放入共享文件夹目录。 请务必将 Web 应用程序本身部署到 Web 服务器。 务必在清单文件的 **SourceLocation** 元素中指定 URL。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. 在 Excel、Word 或 PowerPoint 中，选择功能区上“插入”**** 选项卡中的“我的外接程序”****。

3. 在“Office 外接程序”对话框的顶部，选择“共享文件夹”。 ** ** ** **

4. 依次选择外接程序名称和“确定”****，以插入外接程序。


## <a name="see-also"></a>另请参阅

- [验证并排查清单问题](troubleshoot-manifest.md)
- [发布 Office 外接程序](../publish/publish.md)
    

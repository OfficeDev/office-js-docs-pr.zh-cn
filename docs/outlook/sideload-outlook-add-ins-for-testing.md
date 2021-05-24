---
title: 旁加载 Outlook 外接程序进行测试
description: 利用旁加载来安装 Outlook 外接程序以供测试，无需先将其置于外接程序目录中。
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 9d0fb246f6522c745658a09fce6934ee44d5079a
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555190"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>旁加载 Outlook 外接程序进行测试

可以使用旁加载安装 Outlook 外接程序进行测试，而无需首先将其置于外接程序目录中。

## <a name="sideload-automatically"></a>自动旁加载

如果你使用适用于 Outlook 加载项的[Yeoman](https://github.com/OfficeDev/generator-office)生成器Office加载项，则最好通过命令行进行旁加载。 这将利用我们的工具和通过一个命令跨所有受支持的设备进行旁加载。

1. 使用命令行导航到 Yeoman 生成的加载项项目的根目录。 运行命令 `npm start`。

1. 你的Outlook加载项将自动旁加载Outlook桌面计算机上。 你将看到一个对话框，说明尝试旁加载外接程序，列出清单文件的名称和位置。 选择 **"确定**"，这将注册清单。

    > [!IMPORTANT]
    > 如果清单包含错误或清单路径无效，您将收到错误消息。

1. 如果清单中不包含任何错误且路径有效，外接程序现在将旁加载，并可在桌面和 web Outlook使用。 它还将安装在所有受支持的设备上。

## <a name="sideload-manually"></a>手动旁加载

尽管我们强烈建议通过命令行自动旁加载，如上一节所述，但您也可以基于 Outlook 客户端手动旁加载 Outlook 外接程序。

### <a name="outlook-on-the-web"></a>Outlook 网页版

在 Web 上旁加载加载项Outlook取决于使用的是新版还是经典版。

- 如果邮箱工具栏类似于下图，请参阅[在全新 Outlook 网页版中旁加载外接程序](#new-outlook-on-the-web)。

    ![全新 Outlook 网页版工具栏的部分屏幕截图](../images/outlook-on-the-web-new-toolbar.png)

- 如果邮箱工具栏类似于下图，请参阅[在经典 Outlook 网页版中旁加载外接程序](#classic-outlook-on-the-web)。

    ![经典 Outlook 网页版工具栏的部分屏幕截图](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> 如果你的组织在邮箱工具栏中添加了自己的徽标，则你看到的界面可能会与前面的图像略有不同。

### <a name="new-outlook-on-the-web"></a>Outlook Web 上的新网站

1. 转到 [Outlook 网页版](https://outlook.office.com)。

1. 创建新邮件。

1. 从新邮件的底部选择 **...**，然后从出现的菜单中选择“**获取外接程序**”。

    ![全新 Outlook 网页版中的邮件撰写窗口（突出显示了“获取外接程序”选项）](../images/outlook-on-the-web-new-get-add-ins.png)

1. 在“**Outlook 外接程序**”对话框中，选择“**我的外接程序**”。

    ![全新 Outlook 网页版中的“Outlook 外接程序”对话框（已选中“我的外接程序”）](../images/outlook-on-the-web-new-my-add-ins.png)

1. 在对话框底部找到“**自定义外接程序**”部分。 选择“**添加自定义外接程序**”链接，然后选择“**从文件添加**”。

    ![指向“从文件添加”选项的“管理外接程序”屏幕截图](../images/outlook-sideload-desktop-add-from-file.png)

1. 找到您的自定义外接程序清单文件并进行安装。在安装过程中接受所有提示。

### <a name="classic-outlook-on-the-web"></a>经典Outlook网页

1. 转到 [Outlook 网页版](https://outlook.office.com)。

1. 选择右上部分的齿轮图标，然后选择“**管理外接程序**”。

    ![Outlook 网页版屏幕截图（指向“管理外接程序”选项）](../images/outlook-sideload-web-manage-integrations.png)

1. 在“管理加载项”页中，选择“加载项”，然后选择“我的加载项”。

    ![Outlook 网页版应用商店对话框（已选中“我的外接程序”）](../images/outlook-sideload-store-select-add-ins.png)

1. 在对话框底部找到“**自定义外接程序**”部分。 选择“**添加自定义外接程序**”链接，然后选择“**从文件添加**”。

    ![指向“从文件添加”选项的“管理外接程序”屏幕截图](../images/outlook-sideload-desktop-add-from-file.png)

1. 找到您的自定义外接程序清单文件并进行安装。在安装过程中接受所有提示。

### <a name="outlook-on-the-desktop"></a>Outlook桌面上

#### <a name="outlook-2016-or-later"></a>Outlook 2016或更高版本

1. 在 Outlook 2016 或 Mac 上打开 Windows 或更高版本。

1. 选择功能区上的“**获取外接程序**”按钮。

    ![Outlook 2016"获取外接程序"按钮的自定义功能区](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > 如果在加载项版本中看不到"获取加载项"按钮，Outlook：
    >
    > - **功能** 区上的"存储"按钮（如果可用）。
    >
    >   OR
    >
    > - **"** 文件"菜单，然后选择"信息"选项卡上的"管理外接程序"按钮，以在Web 上的Outlook打开"外接程序"对话框。 <br>有关 Web 体验的更多内容，请参阅上一部分在 Web 上的 Outlook[旁加载外接程序](#outlook-on-the-web)。

1. 如果对话框顶部附近有选项卡，请确保已选中" **加载项** "选项卡。 选择 **"我的外接程序"。**

    ![Outlook 2016 应用商店对话框（已选中“我的外接程序”）](../images/outlook-sideload-store-select-add-ins.png)

1. 在对话框底部找到“自定义加载项”部分。 选择“添加自定义加载项”链接，然后选择“从文件添加”。

    ![“应用商店”屏幕截图（指向“从文件添加”选项）](../images/outlook-sideload-desktop-add-from-file.png)

1. 找到您的自定义外接程序清单文件并进行安装。在安装过程中接受所有提示。

#### <a name="outlook-2013"></a>Outlook 2013

1. 在 Outlook 2013 上Windows。

1. 选择 **"文件**"菜单，然后选择"信息"选项卡上的"管理外接程序"按钮。Outlook浏览器中打开 Web 版本。

1. 按照"在 Web 上旁加载外接程序[Outlook"部分](#outlook-on-the-web)的步骤，具体步骤Outlook Web 上的外接程序版本。

## <a name="remove-a-sideloaded-add-in"></a>删除旁加载的外接程序

在所有版本的 Outlook 中，删除旁加载加载项的关键是列出已安装加载项的"我的加载项"对话框。 选择外接程序 `...` () 省略号，然后选择"删除 **"。**

若要 **导航到** Outlook 客户端的"我的外接程序"对话框，请使用本文前面部分中列出的用于手动旁加载的最后步骤 [](#sideload-manually)。

若要从 Outlook 中删除旁加载的外接程序，请使用本文前面介绍的步骤在列出已安装外接程序的对话框的"自定义外接程序"部分查找外接程序。选择外接程序 () 的省略号，然后选择"删除" `...` 以删除该特定外接程序。  关闭该对话框。

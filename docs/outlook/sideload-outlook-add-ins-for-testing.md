---
title: 旁加载 Outlook 外接程序进行测试
description: 利用旁加载来安装 Outlook 外接程序以供测试，无需先将其置于外接程序目录中。
ms.date: 07/09/2020
localization_priority: Normal
ms.openlocfilehash: 9b44b988ddd6552d5f7d14088a0b6f3ae1e410ed
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093880"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>旁加载 Outlook 外接程序进行测试

可以使用旁加载安装 Outlook 外接程序进行测试，而无需首先将其置于外接程序目录中。

## <a name="sideload-an-add-in-in-outlook-on-the-web"></a>在 Outlook 网页版中旁加载外接程序

在 web 上的 Outlook 中旁加载外接程序的过程取决于您使用的是新版本还是经典版本。

- 如果邮箱工具栏类似于下图，请参阅[在全新 Outlook 网页版中旁加载外接程序](#sideload-an-add-in-in-the-new-outlook-on-the-web)。

    ![全新 Outlook 网页版工具栏的部分屏幕截图](../images/outlook-on-the-web-new-toolbar.png)

- 如果邮箱工具栏类似于下图，请参阅[在经典 Outlook 网页版中旁加载外接程序](#sideload-an-add-in-in-classic-outlook-on-the-web)。

    ![经典 Outlook 网页版工具栏的部分屏幕截图](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> 如果你的组织在邮箱工具栏中添加了自己的徽标，则你看到的界面可能会与前面的图像略有不同。

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a>在全新 Outlook 网页版中旁加载外接程序

1. 转到 [Office 365 中的 Outlook](https://outlook.office.com)。

1. 在 Outlook 网页版中，创建新邮件。

1. 从新邮件的底部选择 **...**，然后从出现的菜单中选择“**获取外接程序**”。

    ![全新 Outlook 网页版中的邮件撰写窗口（突出显示了“获取外接程序”选项）](../images/outlook-on-the-web-new-get-add-ins.png)

1. 在“**Outlook 外接程序**”对话框中，选择“**我的外接程序**”。

    ![全新 Outlook 网页版中的“Outlook 外接程序”对话框（已选中“我的外接程序”）](../images/outlook-on-the-web-new-my-add-ins.png)

1. 在对话框底部找到“**自定义外接程序**”部分。 选择“**添加自定义外接程序**”链接，然后选择“**从文件添加**”。

    ![指向“从文件添加”选项的“管理外接程序”屏幕截图](../images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a>在经典 Outlook 网页版中旁加载外接程序

1. 转到 [Office 365 中的 Outlook](https://outlook.office.com)。

1. 选择右上部分的齿轮图标，然后选择“**管理外接程序**”。

    ![Outlook 网页版屏幕截图（指向“管理外接程序”选项）](../images/outlook-sideload-web-manage-integrations.png)

1. 在“管理加载项”**** 页中，选择“加载项”****，然后选择“我的加载项”****。

    ![Outlook 网页版应用商店对话框（已选中“我的外接程序”）](../images/outlook-sideload-store-select-add-ins.png)

1. 在对话框底部找到“**自定义外接程序**”部分。 选择“**添加自定义外接程序**”链接，然后选择“**从文件添加**”。

    ![指向“从文件添加”选项的“管理外接程序”屏幕截图](../images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a>在 Outlook 桌面版中旁加载外接程序

### <a name="outlook-2016-or-later"></a>Outlook 2016 或更高版本

1. 在 Windows 或 Mac 上打开 Outlook 2016 或更高版本。

1. 选择功能区上的“**获取外接程序**”按钮。

    ![Outlook 2016 功能区（指向“应用商店”按钮）](../images/outlook-sideload-desktop-store.png)

    > [!NOTE]
    > 如果没有在使用的 Outlook 版本中看到“**获取外接程序**”按钮，请改为选择功能区上的“**应用商店**”按钮。

1. 选择“**外接程序**”，然后选择“**我的外接程序**”。

    ![Outlook 2016 应用商店对话框（已选中“我的外接程序”）](../images/outlook-sideload-store-select-add-ins.png)

1. 在对话框底部找到“自定义加载项”**** 部分。 选择“添加自定义加载项”**** 链接，然后选择“从文件添加”****。

    ![“应用商店”屏幕截图（指向“从文件添加”选项）](../images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

### <a name="outlook-2013"></a>Outlook 2013

1. 在 Windows 上打开 Outlook 2013。

1. 选择 "**文件**" 菜单，然后选择 "**信息**" 选项卡上的 "**管理外接程序**" 按钮。 Outlook 将打开浏览器。

1. 按照您的 Outlook 网页版上的[旁加载中的加载](#sideload-an-add-in-in-outlook-on-the-web)项中的步骤，在 web 上的 outlook 的 "web" 部分中执行。

## <a name="remove-a-sideloaded-add-in"></a>删除旁加载加载项

若要从 Outlook 中删除旁加载外接程序，请使用本文中前面所述的步骤，在列出已安装加载项的对话框的 "**自定义外接程序**" 部分中查找该外接程序。选择外接程序的省略号 (`...`) ，然后选择 "**删除**" 以删除该特定外接程序。
---
title: 在 iPad 和 Mac 上旁加载 Office 加载项以供测试
description: ''
ms.date: 02/25/2019
localization_priority: Priority
ms.openlocfilehash: dc0fad24d7f4f062fb0115edcc58a37d8d9052da
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359252"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>在 iPad 和 Mac 上旁加载 Office 加载项以供测试

若要查看你的外接程序在 Office for iOS 中如何运行，可以使用 iTunes 将你的外接程序的清单旁加载到 iPad，或直接将你的外接程序的清单旁加载到 Office for Mac 中。此操作并不能使你在运行时对其设置断点和调试代码，但你可以查看其行为方式，并验证 UI 可用且正确呈现。 

## <a name="prerequisites-for-office-for-ios"></a>Office for iOS 的先决条件

- 安装了 [iTunes](https://www.apple.com/itunes/download/) 的 Windows 或 Mac 计算机。
    
- 安装了 [Excel for iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) 的 iPad（运行 iOS 8.2 或更高版本）以及同步电缆。
    
- 您要测试的外接程序的清单 .xml 文件。
    

## <a name="prerequisites-for-office-for-mac"></a>Office for Mac 的先决条件

- 在已安装 [Office for Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) 的情况下可运行 OS X v10.10 "Yosemite" 或更高版本的 Mac。
    
- Word for Mac 版本 15.18 (160109)。
   
- Excel for Mac 版本 15.19 (160206)。

- PowerPoint for Mac 版本 15.24 (160614)
    
- 你想要测试的外接程序的清单 .xml 文件。
    

## <a name="sideload-an-add-in-on-excel-or-word-for-ipad"></a>将外接程序旁加载到 Excel for iPad 或 Word for iPad 上

1. 使用同步电缆将 iPad 连接到你的计算机。如果是第一次将 iPad 连接到计算机，系统将提示“**信任此计算机？**”。选择“**信任**”继续执行操作。

2. 在 iTunes 中，选择菜单栏下的“**iPad**”图标。

3. 在 iTunes 左侧的"设置"下，选择"应用程序"。

4. 在 iTunes 右侧，向下滚动到"文件共享"，然后在"外接程序"列下选择"Excel"或"Word"。

5. 在"Excel"或"Word 文档"列底部，选择"添加文件"，然后选择您要旁加载的外接程序的清单 .xml 文件。 
    
6. 在你的 iPad 上打开 Excel 或 Word 应用。如果 Excel 或 Word 应用已运行，则选择“**首页**”按钮，然后关闭并重新启动该应用。
    
7. 打开一个文档。
    
8. 选择“**插入**”选项卡上的“**外接程序**”。旁加载的外接程序可在“**外接程序**”UI 中的“**开发人员**”标题下插入。
    
    ![在 Excel 应用程序中插入的加载项](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-on-office-for-mac"></a>在 Office for Mac 上旁加载加载项

> [!NOTE]
> 若要旁加载 Outlook for Mac 加载项，请参阅[旁加载 Outlook 加载项以供测试](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)。

1. 打开“**终端**”并转到以下文件夹之一，你将在其中保存外接程序的清单文件。如果 `wef` 文件夹在你的计算机上不存在，请创建它。
    
    - 对于 Word：`/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`    
    - 对于 Excel：`/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`
    - 对于 PowerPoint：`/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`
    
2. 在“**查找程序**”中使用命令 `open .`（包括句点或点）打开该文件夹。将你的外接程序的清单文件复制到该文件夹中。
    
    ![Office for Mac 中的 Wef 文件夹](../images/all-my-files.png)

3. 打开 Word，然后打开一个文档。如果 Word 已运行，则重新启动它。
    
4. 在 Word 中，选择“**插入**” > “**外接程序**” > “**我的外接程序**”（下拉菜单），然后选择外接程序。
    
    ![Office for Mac 中的“我的加载项”](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > 旁加载的加载项不会显示在“我的加载项”对话框中。它们仅显示在下拉菜单中（单击“插入”**** 选项卡上“我的加载项”右侧的向下小箭头）。旁加载的加载项在此菜单中的“开发人员加载项”**** 标题下列出。 
    
5. 验证加载项是否在 Word 中显示。
    
    ![Office for Mac 中显示的 Office 加载项](../images/lorem-ipsum-wikipedia.png)
    
    > [!NOTE]
    > 出于性能方面的考虑，加载项通常在 Office for Mac 中缓存。 开发加载项时，如果需要强制对其进行重新加载，则可以清除 `Users/<usr>/Library/Containers/com.Microsoft.OsfWebHost/Data/` 文件夹。 如果该文件夹不存在，请清除 `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/` 文件夹中的文件。

## <a name="see-also"></a>另请参阅

- [在 iPad 和 Mac 上调试 Office 加载项](debug-office-add-ins-on-ipad-and-mac.md)

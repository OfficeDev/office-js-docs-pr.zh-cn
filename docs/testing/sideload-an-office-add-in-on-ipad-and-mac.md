---
title: 在 iPad 和 Mac 上旁加载 Office 加载项以供测试
description: 通过旁加载在 iPad 和 Mac 上测试 Office 加载项。
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 22271409cdacd8f3e32039743b8916b1fb87252f
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238069"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>在 iPad 和 Mac 上旁加载 Office 加载项以供测试

若要查看加载项在 iOS 版 Office 中如何运行，可以使用 iTunes 将加载项的清单旁加载到 iPad，或直接将加载项的清单旁加载到 Mac 版 Office 中。此操作并不能使你在运行时对其设置断点和调试代码，但你可以查看其行为方式，并验证 UI 可用且正确呈现。

## <a name="prerequisites-for-office-on-ios"></a>iOS 版 Office 的先决条件

- 安装了 [iTunes](https://www.apple.com/itunes/download/) 的 Windows 或 Mac 计算机。
  > [!IMPORTANT]
  > 如果运行的是 macOS Catalina， [则 iTunes](https://support.apple.com/HT210200) 不再可用，因此您应该按照本文稍后使用 [macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) 在 Excel 或 iPad 上的 Word 旁加载加载项部分中的说明进行操作。

- 运行 iOS 8.2 或更高版本（已安装 [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) 或 [Word）](https://apps.apple.com/app/microsoft-word/id586447913) 的 iPad，以及同步电缆。

- 你想要测试的外接程序的清单 .xml 文件。

## <a name="prerequisites-for-office-on-mac"></a>Mac 版 Office 的先决条件

- 在已安装 [Mac 版 Office](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) 的情况下可运行 OS X v10.10 "Yosemite" 或更高版本的 Mac。

- Mac 版本 15.18 (160109) 上的 Word。

- Mac 版本 15.19 (160206) 上的 Excel。

- Mac 版本 15.24 (160614) 上的 PowerPoint

- 你想要测试的外接程序的清单 .xml 文件。

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>使用 iTunes 在 Excel 或 iPad 上的 Word 上旁加载加载项

1. 使用同步电缆将 iPad 连接到你的计算机。 如果你第一次将 iPad 连接到计算机，系统将提示你信任 **此计算机？。** 选择“**信任**”继续执行操作。

2. 在 iTunes 中，选择菜单栏下的“iPad”图标。

3. 在 iTunes 左侧的“设置”下，选择“应用程序”。

4. 在 iTunes 右侧，向下滚动到“文件共享”，然后在“外接程序”列下选择“Excel”或“Word”。

5. 在 **"Excel"** 或 **"Word 文档**"列的底部，选择"添加文件"，然后选择要旁加载的外接程序的清单 .xml 文件。

6. 在你的 iPad 上打开 Excel 或 Word 应用。 如果 Excel 或 Word 应用已在运行，请选择"开始"按钮，然后关闭并重新启动该应用程序。

7. 打开一个文档。

8. 在"插入"选项卡上选择"外接程序"。 (在"插入"选项卡上，可能需要水平滚动，直到看到"加载项"按钮。) 旁加载的外接程序可用于在外接程序 **UI** 的"开发人员"标题下插入。  

    ![在 Excel 应用程序中插入的加载项](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>使用 macOS Catalina 在 Excel 或 iPad 上的 Word 上旁加载加载项

> [!IMPORTANT]
> 随着 macOS Catalina 的引入，Apple 在 Mac 上废弃 [了 iTunes，](https://support.apple.com/HT210200) 并提供了将应用旁加载到 **Finder 所需的集成功能**。

1. 使用同步电缆将 iPad 连接到你的计算机。 如果你第一次将 iPad 连接到计算机，系统将提示你信任 **此计算机？。** 选择“**信任**”继续执行操作。 此外，还可能会询问你是新 iPad 还是正在还原 iPad。

2. 在 Finder 中的" **位置"下**，选择 **菜单栏下方的 iPad** 图标。

3. 在查找器窗口的顶部，单击 **"文件**"，然后找到 **Excel** 或 **Word。**

4. 从不同的"查找器"窗口中，拖放manifest.xml加载到第一个查找器窗口中的 **Excel** 或 **Word** 文件的外接程序的 manifest.xml 文件。

5. 在你的 iPad 上打开 Excel 或 Word 应用。 如果 Excel 或 Word 应用已在运行，请选择"开始"按钮，然后关闭并重新启动该应用程序。

6. 打开一个文档。

7. 在"插入"选项卡上选择"外接程序"。 (在"插入"选项卡上，可能需要水平滚动，直到看到"加载项"按钮。) 旁加载的外接程序可用于在外接程序 **UI** 的"开发人员"标题下插入。  

    ![在 Excel 应用程序中插入的加载项](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a>在 Mac 版 Office 中旁加载加载项

> [!NOTE]
> 若要旁加载 Mac 版 Outlook 加载项，请参阅[旁加载 Outlook 加载项进行测试](../outlook/sideload-outlook-add-ins-for-testing.md)。

1. 打开 **终端** 并转到以下文件夹之一，你将在其中保存加载项的清单文件。 如果 `wef` 文件夹在你的计算机上不存在，请创建它。

    - 对于 Word：`/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - 对于 Excel：`/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - 对于 PowerPoint：`/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

2. 使用命令命令打开 **Finder** 中的文件夹 (`open .` 包括句点或点) 。 将你的外接程序的清单文件复制到该文件夹中。

    ![Mac 版 Office 中的 Wef 文件夹](../images/all-my-files.png)

3. 打开 Word，然后打开一个文档。如果 Word 已运行，则重新启动它。

4. 在 Word中，选择"在外接程序 (下拉菜单中) 外接程序"，然后选择  >    >  您的外接程序。

    ![Mac 版 Office 中的“我的加载项”](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > 旁加载的加载项不会显示在“我的加载项”对话框中。它们仅显示在下拉菜单中（单击“插入”选项卡上“我的加载项”右侧的向下小箭头）。旁加载的加载项在此菜单中的“开发人员加载项”标题下列出。

5. 验证加载项是否在 Word 中显示。

    ![Mac 版 Office 中显示的 Office 加载项](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>删除旁加载的加载项

可以通过清除您计算机的 Office 缓存来删除以前旁加载的外接程序。 有关如何清除每个平台和应用程序缓存的详细信息，请参阅"清除 Office 缓存"一 [文](clear-cache.md)。

## <a name="see-also"></a>另请参阅

- [在 iPad 和 Mac 上调试 Office 加载项](debug-office-add-ins-on-ipad-and-mac.md)

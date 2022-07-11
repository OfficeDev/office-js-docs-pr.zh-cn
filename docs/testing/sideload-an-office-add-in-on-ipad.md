---
title: 在 iPad 上旁加载 Office 加载项以进行测试
description: 通过旁加载在 iPad 上测试 Office 加载项。
ms.date: 06/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0ba52ae78bed36c4eb8130c714577a1b0899aeb6
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713199"
---
# <a name="sideload-office-add-ins-on-ipad-for-testing"></a>在 iPad 上旁加载 Office 加载项以进行测试

若要了解外接程序如何在 iOS 上的 Office 中运行，可以使用 iTunes 将外接程序的清单旁加载到 iPad 上。 此操作并不能使你在运行时对其设置断点和调试代码，但你可以查看其行为方式，并验证 UI 可用且正确呈现。

> [!NOTE]
> 若要旁加载 Outlook 外接程序，请参阅[旁加载 Outlook 外接程序进行测试](../outlook/sideload-outlook-add-ins-for-testing.md)。

## <a name="prerequisites-for-office-on-ios"></a>iOS 版 Office 的先决条件

- 安装了 [iTunes](https://www.apple.com/itunes/download/) 的 Windows 或 Mac 计算机。
  > [!IMPORTANT]
  > 如果运行的是 macOS Catalina， [则 iTunes 不再可用](https://support.apple.com/HT210200) ，因此应按照本文稍后在 [Excel 或 iPad 上的 Word 上使用 macOS Catalina 旁加载加](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) 载项部分中的说明操作。

- 运行 iOS 8.2 或更高版本且安装了 [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) 或 [Word](https://apps.apple.com/app/microsoft-word/id586447913) 的 iPad 以及同步电缆。

- 你想要测试的外接程序的清单 .xml 文件。

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>使用 iTunes 在 Excel 或 iPad 上的 Word 上旁加载加载项

1. 使用同步电缆将 iPad 连接到你的计算机。 如果是第一次将 iPad 连接到计算机，系统会提示你 **使用“信任此计算机？”**。 选择“**信任**”继续执行操作。

2. 在 iTunes 中，选择菜单栏下的“iPad”图标。

3. 在 iTunes 左侧的“设置”下，选择“应用程序”。

4. 在 iTunes 右侧，向下滚动到“文件共享”，然后在“外接程序”列下选择“Excel”或“Word”。

5. 在 **Excel** 或 **Word Documents** 列的底部，选择 **“添加文件**”，然后选择要旁加载的加载项的清单.xml文件。

6. 在你的 iPad 上打开 Excel 或 Word 应用。 如果 Excel 或 Word 应用已在运行，请选择 **“开始** ”按钮，然后关闭并重启应用。

7. 打开一个文档。

8. 在“**插入**”选项卡上选择 **加** 载项。 (“**插入**”选项卡上，可能需要水平滚动，直到看到 **“加** 载项”按钮。) 旁加载的外接程序可用于在 **加载项 UI 的****“开发人员**”标题下插入。

    ![在 Excel 应用中插入加载项。](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>使用 macOS Catalina 在 Excel 或 iPad 上的 Word 上旁加载加载项

> [!IMPORTANT]
> 随着 macOS Catalina 的推出， [Apple 在 Mac 上停用了 iTunes](https://support.apple.com/HT210200) ，并集成了将应用旁加载到 **Finder 中** 所需的功能。

1. 使用同步电缆将 iPad 连接到你的计算机。 如果是第一次将 iPad 连接到计算机，系统会提示你 **使用“信任此计算机？”**。 选择“**信任**”继续执行操作。 还可能会询问这是新的 iPad 还是正在还原 iPad。

2. 在 Finder 的 **“位置**”下，选择菜单栏下方的 **iPad** 图标。

3. 在“查找器”窗口顶部，单击 **“文件**”，然后找到 **Excel** 或 **Word**。

4. 从其他 Finder 窗口中，拖放要在第一个查找器窗口中将加载项的manifest.xml文件旁加载到 **Excel** 或 **Word** 文件上。

5. 在你的 iPad 上打开 Excel 或 Word 应用。 如果 Excel 或 Word 应用已在运行，请选择 **“开始** ”按钮，然后关闭并重启应用。

6. 打开一个文档。

7. 在“**插入**”选项卡上选择 **加** 载项。 (“**插入**”选项卡上，可能需要水平滚动，直到看到 **“加** 载项”按钮。) 旁加载的外接程序可用于在 **加载项 UI 的****“开发人员**”标题下插入。

    ![在 Excel 应用中插入加载项。](../images/excel-insert-add-in.png)

## <a name="remove-a-sideloaded-add-in"></a>删除旁加载的加载项

可以通过清除计算机上的 Office 缓存来删除以前旁加载的加载项。 有关如何清除每个平台和应用程序的缓存的详细信息，请参阅“ [清除 Office 缓存](clear-cache.md)”一文。

## <a name="see-also"></a>另请参阅

- [在 Mac 上旁加载 Office 加载项以进行测试](sideload-an-office-add-in-on-mac.md)
- [在 Mac 上调试 Office 加载项](debug-office-add-ins-on-ipad-and-mac.md)
- [旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)

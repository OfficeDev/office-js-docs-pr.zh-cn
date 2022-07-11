---
title: 在 Mac 上旁加载 Office 加载项以进行测试
description: 通过旁加载在 Mac 上测试 Office 加载项。
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 38ed5f5dba2d379b6137a098240021bd642d6e11
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713205"
---
# <a name="sideload-office-add-ins-on-mac-for-testing"></a>在 Mac 上旁加载 Office 加载项以进行测试

若要了解加载项在 Mac 上的 Office 上如何运行，可以旁加载外接程序的清单。 此操作并不能使你在运行时对其设置断点和调试代码，但你可以查看其行为方式，并验证 UI 可用且正确呈现。

> [!NOTE]
> 若要旁加载 Outlook 外接程序，请参阅[旁加载 Outlook 外接程序进行测试](../outlook/sideload-outlook-add-ins-for-testing.md)。

## <a name="prerequisites-for-office-on-mac"></a>Mac 版 Office 的先决条件

- 在已安装 [Mac 版 Office](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) 的情况下可运行 OS X v10.10 "Yosemite" 或更高版本的 Mac。

- Mac 版本 15.18 (160109) 上的 Word。

- Mac 版本 15.19 (160206) 上的 Excel。

- PowerPoint on Mac 版本 15.24 (160614) 。

- 你想要测试的外接程序的清单 .xml 文件。

## <a name="sideload-an-add-in-in-office-on-mac"></a>在 Mac 版 Office 中旁加载加载项

1. 使用 **Finder** 旁加载清单文件。 打开 **Finder** ，然后输入 Command+Shift+G 以打开 **“转到文件夹** ”对话框。

1. 根据要用于旁加载的应用程序，输入以下文件路径之一。 如果 `wef` 文件夹在你的计算机上不存在，请创建它。

    - 对于 Word：`/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - 对于 Excel：`/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - 对于 PowerPoint：`/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

        > [!NOTE]
        > 其余步骤介绍如何旁加载 Word 加载项。

1. 将加载项的清单文件复制到此 `wef` 文件夹。

    ![Office on Mac 中的 Wef 文件夹。](../images/all-my-files.png)

1. 打开 Word，然后打开一个文档。 如果它已经在运行，请重启 Word。

1. 在 Word 中，选择“ **插入** > **加载项** > **我的加载项** ” (下拉菜单) ，然后选择加载项。

    ![我的 Office on Mac 加载项。](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > 旁加载的加载项不会显示在“我的加载项”对话框中。它们仅显示在下拉菜单中（单击“插入”选项卡上“我的加载项”右侧的向下小箭头）。旁加载的加载项在此菜单中的“开发人员加载项”标题下列出。

1. 验证加载项是否在 Word 中显示。

    ![Office 加载项显示在 Mac 上的 Office 中。](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>删除旁加载的加载项

可以通过清除计算机上的 Office 缓存来删除以前旁加载的加载项。 有关如何清除每个平台和应用程序的缓存的详细信息，请参阅“ [清除 Office 缓存](clear-cache.md)”一文。

## <a name="see-also"></a>另请参阅

- [在 iPad 上旁加载 Office 加载项以进行测试](sideload-an-office-add-in-on-ipad.md)
- [在 Mac 上调试 Office 加载项](debug-office-add-ins-on-ipad-and-mac.md)
- [旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)

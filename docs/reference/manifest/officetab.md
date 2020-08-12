---
title: 清单文件中的 OfficeTab 元素
description: OfficeTab 元素定义在其中显示外接程序命令的功能区选项卡。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 9b07ce1e57329e796545610e0c61a2c11d1ed55d
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641437"
---
# <a name="officetab-element"></a>OfficeTab 元素

定义在其上显示外接程序命令的功能区选项卡。 这可以是默认选项卡 ("**主页**"、"**邮件**" 或 "**会议**) "，也可以是由加载项定义的自定义选项卡。 此元素是必需的。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  组      | 是 |  定义一组命令。对于每个外接程序，只能将一个组添加到默认选项卡。  |

下面是主机的有效选项卡 `id` 值。 以**粗体显示**的值在桌面和联机 (中均受支持（例如，word 2016 或更高版本位于 web 上的 Windows 和 word) 。

### <a name="outlook"></a>Outlook

- **TabDefault**

### <a name="word"></a>Word

- **TabHome**
- **TabInsert**
- TabWordDesign
- **TabPageLayoutWord**
- TabReferences
- TabMailings
- TabReviewWord
- **TabView**
- TabDeveloper
- TabAddIns
- TabBlogPost
- TabBlogInsert
- TabPrintPreview
- TabOutlining
- TabConflicts
- TabBackgroundRemoval
- TabBroadcastPresentation

### <a name="excel"></a>Excel

- **TabHome**
- **TabInsert**
- TabPageLayoutExcel
- TabFormulas
- **TabData**
- **TabReview**
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabBackgroundRemoval

### <a name="powerpoint"></a>PowerPoint

- **TabHome**
- **TabInsert**
- **TabDesign**
- **TabTransitions**
- **TabAnimations**
- TabSlideShow
- TabReview
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabMerge
- TabGrayscale
- TabBlackAndWhite
- TabBroadcastPresentation
- TabSlideMaster
- TabHandoutMaster
- TabNotesMaster
- TabBackgroundRemoval
- TabSlideMasterHome

### <a name="onenote"></a>OneNote

- **TabHome**
- **TabInsert**
- **TabView**
- TabDeveloper
- TabAddIns

## <a name="group"></a>组

选项卡中的一组 UI 扩展点。 **Id**属性是必需的，并且每个**id**在清单中必须是唯一的。 **Id**是最多为125个字符的字符串。 查看 [Group 元素](group.md)。

## <a name="officetab-example"></a>OfficeTab 示例

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```

---
title: 清单文件中的 OfficeTab 元素
description: OfficeTab 元素定义显示外接程序命令的功能区选项卡。
ms.date: 06/20/2019
ms.localizationpriority: medium
ms.openlocfilehash: 94af0e8744c496538a506fdc4626cd3f515129bb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152333"
---
# <a name="officetab-element"></a>OfficeTab 元素

定义在其上显示外接程序命令的功能区选项卡。 这可以是默认选项卡 **(、消息** 或会议) ，也可以是由外接程序定义的自定义选项卡。 此元素是必需的。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  组      | 是 |  定义一组命令。对于每个外接程序，只能将一个组添加到默认选项卡。  |

以下是应用程序的有效选项卡 `id` 值。 粗体值 **在** 桌面和联机设备中均受支持 (例如，Word 2016或更高版本Windows和Word web 版) 。

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

选项卡中的一组 UI 扩展点。 **id** 属性是必需的，并且每个 **id** 在清单中必须是唯一的。 **id** 是最多包含 125 个字符的字符串。 查看 [Group 元素](group.md)。

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

---
title: 清单文件中的 Control 元素
description: 定义用于执行操作或启动任务窗格的控件。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa7ff9b0162070b378352ce187de15a34323b998
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467834"
---
# <a name="control-element"></a>Control 元素

定义用于执行操作或启动任务窗格的控件。 **Control** 元素可以是按钮选项，也可以是菜单选项。 **Group** 元素中至少需包括一个 [Control](group.md)。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) (对于任务窗格外接程序.) 
- 某些子元素可能与其他要求集相关联。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|**xsi:type**|是|正被定义的控件类型。 可以是 、`Button``Menu`、 或 `MobileButton`。 |
|**id**|是|控件元素的 ID。 最多可包含 125 个字符。 在清单中所有 **Control** 元素中必须是唯一的。|

> [!NOTE]
> 在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。 它只适用于 [MobileFormFactor](mobileformfactor.md) 元素内包含的 **Control** 元素。

## <a name="child-elements"></a>子元素

有效的子元素取决于 **xsi：type** 属性的值。

- [Control 元素的按钮类型](control-button.md)
- [Control 元素的菜单类型](control-menu.md)
- [Control 元素的 MobileButton 类型](control-mobilebutton.md)

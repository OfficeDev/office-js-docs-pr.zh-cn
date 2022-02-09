---
title: 清单文件中 Item 元素
description: 指定菜单中的项。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: cd46b46e1466b8cb9bab7e283ddca437721e762e
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467882"
---
# <a name="item-element"></a>Item 元素

指定菜单中的项。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- 当父 **VersionOverrides** 的类型为 Taskpane [1.0 时，AddinCommands](../requirement-sets/add-in-commands-requirement-sets.md) 1.1。
- [当父](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) **VersionOverrides** 类型为 Mail 1.0 时，邮箱 1.3。
- [当父](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) **VersionOverrides** 类型为 Mail 1.1 时，邮箱 1.5。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Label](#label)     | 是 |  按钮文本。 |
|  [Supertip](supertip.md)  | 是 |  按钮的 supertip。    |
|  [Icon](icon.md)      | 是 |  按钮的图像。         |
|  [Action](action.md)    | 是 |  指定要执行的操作。 Item 元素只能有 **一个 Action** **子元素** 。  |
|  [Enabled](enabled.md)    | 否 |  指定加载项启动时是否启用控件。  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | 否 |  指定该按钮是否应该显示在支持自定义上下文选项卡的应用程序和平台组合上。 如果使用，则它必须是第 *一个子* 元素。 |

### <a name="label"></a>标签

通过按钮的唯一属性 **resid** 指定按钮的文本，该属性不能超过 32 个字符，并且必须设置为 [](resources.md) **ShortStrings 元素的 ShortStrings** 子元素中 **String** 元素的 **id** 属性的值。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- 当父 **VersionOverrides** 的类型为 Taskpane [1.0 时，AddinCommands](../requirement-sets/add-in-commands-requirement-sets.md) 1.1。
- [当父](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) **VersionOverrides** 类型为 Mail 1.0 时，邮箱 1.3。
- [当父](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) **VersionOverrides** 类型为 Mail 1.1 时，邮箱 1.5。

## <a name="examples"></a>示例

有关示例，请参阅 [Menu 类型的控件](control-menu.md)。
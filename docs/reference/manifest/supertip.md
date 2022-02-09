---
title: 清单文件中的 Supertip 元素
description: Supertip 元素定义一个丰富的工具提示 (标题和说明) 。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: aab7ab3f17e772940403e75796346020b2b9aebe
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467855"
---
# <a name="supertip"></a>Supertip

定义丰富的工具提示（标题和说明）。 它同时由 [Button 控件和](control-button.md) [Menu 控件使用](control-menu.md)。

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
| [标题](#title) | 是 | supertip 的文本。 |
| [说明](#description) | 是 | supertip 的说明。<br>**注意**： (Outlook) 仅Windows和 Mac 客户端。 |

### <a name="title"></a>标题

必需。 supertip 的文本。 **resid** 属性的长度不能超过 32 个字符，并且必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。

### <a name="description"></a>说明

必需。 supertip 的说明。 **resid** 属性的长度不能超过 32 个字符，并且必须设置为 **LongStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。

> [!NOTE]
> 对于Outlook，只有 Windows 和 Mac 客户端支持 **Description** 元素。

## <a name="example"></a>示例

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

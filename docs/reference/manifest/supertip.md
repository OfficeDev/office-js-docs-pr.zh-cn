---
title: 清单文件中的 Supertip 元素
description: Supertip 元素定义一个丰富的工具提示 (标题和说明) 。
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 5e8b3850d99f6791726b1b2f0545c5fb4b52c554
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771296"
---
# <a name="supertip"></a>Supertip

定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
| [标题](#title) | 是 | supertip 的文本。 |
| [说明](#description) | 是 | supertip 的说明。<br>**注意**： (Outlook) 仅支持 Windows 和 Mac 客户端。 |

### <a name="title"></a>标题

必需。 supertip 的文本。 **resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 ShortStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)

### <a name="description"></a>说明

必需。 supertip 的说明。 **resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 LongStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)

> [!NOTE]
> 对于 Outlook，只有 Windows 和 Mac 客户端支持 **Description** 元素。

## <a name="example"></a>示例

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

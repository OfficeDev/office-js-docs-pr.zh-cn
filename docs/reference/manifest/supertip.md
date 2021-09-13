---
title: 清单文件中的 Supertip 元素
description: Supertip 元素定义一个丰富的工具提示 (标题和说明) 。
ms.date: 05/07/2019
ms.localizationpriority: medium
ms.openlocfilehash: 6c1e73b0aba5923992fba03b78744ae5d34fb5da
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152504"
---
# <a name="supertip"></a>Supertip

定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
| [标题](#title) | 是 | supertip 的文本。 |
| [说明](#description) | 是 | supertip 的说明。<br>**注意**： (Outlook) 仅Windows和 Mac 客户端。 |

### <a name="title"></a>标题

必需。 supertip 的文本。 **resid** 属性的长度不能超过 32 个字符，并且必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md)元素）中 **String** 元素的 **id** 属性的值。

### <a name="description"></a>说明

必需。 supertip 的说明。 **resid** 属性的长度不能超过 32 个字符，并且必须设置为 **LongStrings** 元素（位于 [Resources](resources.md)元素）中 **String** 元素的 **id** 属性的值。

> [!NOTE]
> 对于Outlook，只有 Windows 和 Mac 客户端支持 **Description** 元素。

## <a name="example"></a>示例

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

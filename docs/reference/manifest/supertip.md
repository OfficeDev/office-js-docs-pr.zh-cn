---
title: 清单文件中的 Supertip 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: bae997eda8e1055c5be76382456ba83acca7b91c
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433668"
---
# <a name="supertip"></a>Supertip

定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Title](#title)        | 是 |   supertip 的文本。         |
|  [说明](#description)  | 是 |  supertip 的说明。    |

### <a name="title"></a>标题

必需。SuperTip 的文本。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。

### <a name="description"></a>说明

必需。SuperTip 的描述。 **resid** 属性必须设置为 **LongStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。

## <a name="example"></a>示例

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

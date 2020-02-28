---
title: 清单文件中的 Supertip 元素
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: ab280ec550a58f85082c36a24f5f7c3b4112a214
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325232"
---
# <a name="supertip"></a>Supertip

定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
| [标题](#title) | 是 | supertip 的文本。 |
| [说明](#description) | 是 | supertip 的说明。<br>**注意**：（Outlook）仅支持 Windows 和 Mac 客户端。 |

### <a name="title"></a>Title

必填。 SuperTip 的文本。 **Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。

### <a name="description"></a>说明

必需。 SuperTip 的描述。 **Resid**属性必须设置为[Resources](resources.md)元素中的**LongStrings**元素中**String**元素的**id**属性的值。

> [!NOTE]
> 对于 Outlook，只有 Windows 和 Mac 客户端支持**Description**元素。

## <a name="example"></a>示例

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

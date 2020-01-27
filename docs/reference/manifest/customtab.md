---
title: 清单文件中的 CustomTab 元素
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: c48e526534a3c1295e9c3f0c6fc626df94a874d3
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554011"
---
# <a name="customtab-element"></a>CustomTab 元素

在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。这可以位于默认的选项卡（“**开始**”、“**消息**”或“**会议**”）上，或位于由外接程序定义的自定义选项卡上。

在自定义选项卡上，外接程序可以创建最多 10 个组。每个组限制为 6 个控件，不论它显示在哪个选项卡上。外接程序限定到一个自定义选项卡。

**id** 属性在清单中必须是唯一的。

> [!IMPORTANT]
> 在 Mac 上的 Outlook 中`CustomTab` ，该元素不可用，因此您必须改用[OfficeTab](officetab.md) 。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Group](group.md)      | 是 |  定义一组命令。  |
|  [Label](#label-tab)      | 是 |  CustomTab 或组的标签。  |

### <a name="group"></a>组

必需。查看 [Group 元素](group.md)。

### <a name="label-tab"></a>标签（选项卡）

必需。自定义选项卡的标签。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。


## <a name="customtab-example"></a>CustomTab 示例

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

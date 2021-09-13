---
title: 清单文件中的 AllFormFactors 元素
description: 指定加载项的所有外观设置。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: b579612d73216fab6141501e1c969fb6be1e8495
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151933"
---
# <a name="allformfactors-element"></a>AllFormFactors 元素

指定加载项的所有外观设置。 目前，使用 **AllFormFactors** 的唯一功能是自定义函数。 使用自定义函数时，**AllFormFactors** 是必备元素。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  是 |  定义加载项用于公开功能的位置。 |

## <a name="allformfactors-example"></a>AllFormFactors 示例

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```

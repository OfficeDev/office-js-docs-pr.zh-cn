---
title: 清单文件中的 AllFormFactors 元素
description: 指定加载项的所有外观设置。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa15eb48ec8d3fde125973efcea36067f7cdac39
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340405"
---
# <a name="allformfactors-element"></a>AllFormFactors 元素

指定加载项的所有外观设置。 目前，使用 **AllFormFactors** 的唯一功能是自定义函数。 使用自定义函数时，**AllFormFactors** 是必备元素。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

> [!NOTE]
> 此元素仅在 Excel、Windows Mac 和 Web 上受支持。 它在其他应用程序或 iOS Office Android 上不受支持。

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

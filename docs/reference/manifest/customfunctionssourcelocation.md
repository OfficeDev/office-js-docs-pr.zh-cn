---
title: 清单文件中自定义函数的 SourceLocation 元素
description: 定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="sourcelocation-element-custom-functions"></a>SourceLocation 元素 (自定义函数) 

定义自定义函数在自定义元素中使用的 **Script** 或 **Page** 元素所需的资源Excel。

> [!IMPORTANT]
> 本文仅引用 **作为 Page 或** Script 元素的子元素的 **SourceLocation**。 有关 [基本清单](sourcelocation.md) 的 **SourceLocation** 元素的信息，请参阅 SourceLocation。

**外接程序类型：** 自定义函数

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>属性

| 属性 | 必需 | 说明                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | 是      | 清单的 **Resources** 部分中所定义的 URL 资源的名称。 不能超过 32 个字符。 |

## <a name="child-elements"></a>子元素

无

## <a name="example"></a>示例

```xml
<SourceLocation resid="pageURL"/>
```

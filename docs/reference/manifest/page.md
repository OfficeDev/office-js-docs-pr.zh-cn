---
title: 清单文件中的 Page 元素
description: Page 元素定义自定义函数在自定义页面中使用的 HTML Excel。
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="page-element"></a>Page 元素

定义 Excel 中的自定义函数所使用的 HTML 页面设置。

**外接程序类型：** 自定义函数

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md) 

## <a name="attributes"></a>属性

无

## <a name="child-elements"></a>子元素

|  元素  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  是  | 包含自定义函数所使用的 HTML 文件的资源 ID 的字符串。 |

## <a name="example"></a>示例

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```

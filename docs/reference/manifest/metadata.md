---
title: 清单文件中的 Metadata 元素
description: Metadata 元素定义自定义函数在元数据Excel。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52938155442bb5424a170634d1324de77de2b788
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855532"
---
# <a name="metadata-element"></a>Metadata 元素

定义 Excel 中的自定义函数所使用的元数据设置。

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
|  [SourceLocation](customfunctionssourcelocation.md)  |  是  | 包含自定义函数所使用的 JSON 文件的资源 ID 的字符串。 |

## <a name="example"></a>示例

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```

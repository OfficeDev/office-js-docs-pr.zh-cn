---
title: 清单文件中的 TargetDialect 元素
description: TargetDialect 元素定义此字典支持的区域语言，表示为区域性名称字符串。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ba5c43b6471f11d7599da8542c30618ea1de78e0
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720328"
---
# <a name="targetdialect-element"></a>TargetDialect 元素

定义此字典支持的、表示为区域性名称字符串的区域语言。

**外接程序类型：** 任务窗格

## <a name="syntax"></a>语法

```XML
<TargetDialect>
   string 
</TargetDialect>
```

## <a name="contained-in"></a>包含于

[TargetDialects](targetdialects.md)

## <a name="remarks"></a>注解

指定值采用 BCP 47 语言标记格式，如 `en-US`。

## <a name="see-also"></a>另请参阅

- [创建字典任务窗格外接程序](../../word/dictionary-task-pane-add-ins.md)

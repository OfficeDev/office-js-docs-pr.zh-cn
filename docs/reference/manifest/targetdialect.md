---
title: 清单文件中的 TargetDialect 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8ee97d0851c82bcd8763152a6d0cf4331e0f0bdb
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596870"
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

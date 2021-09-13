---
title: 清单文件中的 TargetDialect 元素
description: TargetDialect 元素定义此字典支持的区域语言，表示为区域性名称字符串。
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: a208b80f1a715c5ee3626f632fb757f347bdcc0a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149432"
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

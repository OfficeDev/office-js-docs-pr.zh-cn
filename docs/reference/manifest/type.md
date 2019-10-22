---
title: 清单文件中的 Type 元素
description: ''
ms.date: 05/03/2019
localization_priority: Normal
ms.openlocfilehash: 1c053d65c5e3c6ce597c9912ec608e0b36bc623b
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/21/2019
ms.locfileid: "33628226"
---
# <a name="type-element"></a>Type 元素

指定等效加载项是 COM 外接程序还是 XLL。

**外接类型：** 任务窗格，自定义函数

## <a name="syntax"></a>语法

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>包含于

[EquivalentAdd-in](equivalentaddin.md)

## <a name="add-in-type-values"></a>外接类型值

必须为`Type`元素指定下列值之一。

- COM：指定等效的加载项是 COM 加载项。
- XLL：指定等效的外接程序是 Excel XLL。

## <a name="see-also"></a>另请参阅

- [让自定义功能与 XLL 用户定义的功能兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [使 Excel 外接程序与现有 COM 外接程序兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
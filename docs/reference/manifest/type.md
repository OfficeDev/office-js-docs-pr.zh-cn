---
title: 清单文件中的 Type 元素
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 28514e25d7877c0452fbf006a31f078cd980d819
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356857"
---
# <a name="type-element"></a>Type 元素

指定等效加载项是 COM 外接程序还是 XLL。

**外接类型:** 任务窗格, 自定义函数

## <a name="syntax"></a>语法

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>包含于

[EquivalentAdd-in](equivalentaddin.md)

## <a name="add-in-type-values"></a>外接类型值

必须为`Type`元素指定下列值之一。

- com: 指定等效的加载项是 COM 加载项。
- XLL: 指定等效的外接程序是 Excel XLL。

## <a name="see-also"></a>另请参阅

- [使自定义函数与 XLL 用户定义的函数兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [使您的 Office 外接程序与现有的 COM 外接程序兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
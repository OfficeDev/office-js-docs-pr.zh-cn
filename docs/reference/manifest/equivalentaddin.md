---
title: 清单文件中的 EquivalentAddin 元素
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 9cb1bb6d7a9cc3df3f4e39f8180b38d47d0a6882
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356848"
---
# <a name="equivalentaddin-element"></a>EquivalentAddin 元素

为等效的 COM 外接程序或 XLL 指定向后兼容性。

**外接类型:** 任务窗格, 自定义函数

## <a name="syntax"></a>语法

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>包含于

[EquivalentAdd-ins](equivalentaddins.md)

## <a name="must-contain"></a>必须包含

[Type](type.md)

## <a name="can-contain"></a>可以包含

[ProgID](progid.md)
[文件名](filename.md)

## <a name="remarks"></a>说明

若要将 COM 加载项指定为等效的`ProgID`加载项, 请同时提供和`Type`元素。 若要将 XLL 指定为等效的外接程序, 请同时`FileName`提供`Type`和元素。

## <a name="see-also"></a>另请参阅

- [使自定义函数与 XLL 用户定义的函数兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [使您的 Office 外接程序与现有的 COM 外接程序兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
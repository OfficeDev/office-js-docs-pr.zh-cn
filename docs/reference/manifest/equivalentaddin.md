---
title: 清单文件中 EquivalentAddin 元素
description: 指定等效 COM 加载项或 XLL 的向后兼容性。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 6404ad561a14a451e4685cc23be930b7ba612e85d1b37e78aa45f9366becf3bc
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57085757"
---
# <a name="equivalentaddin-element"></a>EquivalentAddin 元素

指定等效 COM 加载项或 XLL 的向后兼容性。

**外接程序类型：** 任务窗格、自定义函数

## <a name="syntax"></a>语法

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>包含于

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>必须包含

[类型](type.md)

## <a name="can-contain"></a>可以包含

[ProgId](progid.md) 
[FileName](filename.md)

## <a name="remarks"></a>备注

若要将 COM 加载项指定为等效加载项，请提供 和 `ProgId` `Type` 元素。 若要将 XLL 指定为等效的外接程序，请提供 和 `FileName` `Type` 元素。

## <a name="see-also"></a>另请参阅

- [让自定义功能与 XLL 用户定义的功能兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [让 Office 加载项与现有 COM 加载项兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
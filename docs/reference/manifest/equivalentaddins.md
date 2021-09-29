---
title: 清单文件中 EquivalentAddins 元素
description: 指定与等效 COM 加载项和/或 XLL 的向后兼容性。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: d32f67f49d334a75433aec2d079b45a44a04121a
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990808"
---
# <a name="equivalentaddins-element"></a>EquivalentAddins 元素

指定与等效 COM 加载项和/或 XLL 的向后兼容性。

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**外接程序类型：** 任务窗格、邮件、自定义函数

## <a name="syntax"></a>语法

```XML
<EquivalentAddins>
...  
</EquivalentAddins>  
```

## <a name="contained-in"></a>包含于

[VersionOverrides](versionoverrides.md)

## <a name="must-contain"></a>必须包含

[EquivalentAddin](equivalentaddin.md)

## <a name="see-also"></a>另请参阅

- [让自定义功能与 XLL 用户定义的功能兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [让 Office 加载项与现有 COM 加载项兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
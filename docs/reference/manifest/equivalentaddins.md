---
title: 清单文件中 EquivalentAddins 元素
description: 指定与等效 COM 加载项和/或 XLL 的向后兼容性。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 48f3ef86f71ad3d4f0c759df4583af4cd95e5c5a
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042152"
---
# <a name="equivalentaddins-element"></a>EquivalentAddins 元素

指定与等效 COM 加载项和/或 XLL 的向后兼容性。

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**外接程序类型：** 任务窗格、邮件、自定义函数

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

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
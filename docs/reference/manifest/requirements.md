---
title: 清单文件中的 Requirements 元素
description: "\"要求\" 元素指定要激活的 Office 外接程序所需的最低要求集和方法。"
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 586f05ec68257462cb64a96abf2a34eb31861a5c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611713"
---
# <a name="requirements-element"></a>Requirements 元素

指定 Office 外接程序需要激活的最小 Office JavaScript API 要求（[要求集](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets)和/或方法）集。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[方法](methods.md)|x||x|

## <a name="remarks"></a>注释

有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

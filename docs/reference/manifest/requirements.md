---
title: 清单文件中的 Requirements 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 364ab7c943895e1acecedba7970e54da331a2e6f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870364"
---
# <a name="requirements-element"></a>Requirements 元素

指定适用于 Office 的 JavaScript API 要求（[要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)和/或方法）的最小集，Office 外接程序需要该集才能激活。

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

有关要求集的详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。


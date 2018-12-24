---
title: 清单文件中的 Requirements 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 2544e9b01b2d4d3ddc0a0c6238b4a5b0e6c4f832
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432702"
---
# <a name="requirements-element"></a>Requirements 元素

指定适用于 Office 的 JavaScript API 要求（[要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)和/或方法）的最小集，Office 外接程序需要该集才能激活。

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

|**元素**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Methods](methods.md)|x||x|

## <a name="remarks"></a>注释

有关要求集的详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。


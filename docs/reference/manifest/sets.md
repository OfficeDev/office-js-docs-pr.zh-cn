---
title: 清单文件中的 Sets 元素
description: Set 元素指定 Office 外接程序在激活时所需的最小 Office JavaScript API 集。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bd8f8311bb06a8e9e98fc408aece6395ab5643b1
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641422"
---
# <a name="sets-element"></a>Sets 元素

指定 Office JavaScript API 的最小子集，Office 外接程序需要这些 API 才能激活。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>包含于

[要求](requirements.md)

## <a name="can-contain"></a>可以包含

[Set](set.md)

## <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|字符串|可选|指定所有子[集](set.md)元素的默认**MinVersion**属性值。 默认值为“1.1”。|

## <a name="remarks"></a>注释

有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

有关**Set**元素的**MinVersion**属性和**Sets**元素的**DefaultMinVersion**属性的详细信息，请参阅[在清单中设置需求元素](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。


---
title: 清单文件中的 Set 元素
description: Set 元素指定 Office外接程序所需的 Office JavaScript API 要求集，以便由 Office 激活或替代基本清单设置。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 55e1b25765bfbe53108bc9201c0c851c6ef9161d
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222232"
---
# <a name="set-element"></a>Set 元素

此元素的含义取决于它在清单中的使用位置。

## <a name="in-the-base-manifest"></a>在基本清单中

在基本清单 (即，当您在基本清单中使用的元素是 [OfficeApp](officeapp.md)) 的直接子元素时 **，Set** 元素会从 Office JavaScript API 中指定您的 Office 外接程序需要该要求集，以便 Office 激活。 [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)

**外接程序类型：** 内容、任务窗格、邮件

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>作为 VersionOverrides 元素的子级

指定 Office[](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)版本和平台 (（如 Windows、Mac、Web 和 iOS 或 iPad) ）必须支持的 Office JavaScript API 要求集[，VersionOverrides](versionoverrides.md)才能生效。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 与上一个 [Requirements](requirements.md) 元素相同。

**与以下要求集相关联**：

- 与上一个 [Requirements](requirements.md) 元素相同。

## <a name="syntax"></a>语法

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>包含于

[Sets](sets.md)

## <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|Name|字符串|必需|[要求集](../../develop/office-versions-and-requirement-sets.md)名称。|
|MinVersion|字符串|可选|指定您的外接程序所需的 API 集的最低版本。 替代 **DefaultMinVersion** 的值（如果在父 Sets 元素 [中](sets.md) 指定）。|

## <a name="remarks"></a>注释

有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

有关 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性详细信息，请参阅指定哪些 Office 版本和平台可以托管 [您的外接程序](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。


---
title: 清单文件中的 HighResolutionIconUrl 元素
description: 指定用于表示插入 UX 中的 Office 外接程序和高 DPI 屏幕上的 Office 应用商店的图像的 URL。
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 78a9296f38a688073e516fb78a77bb4cdac822c4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718137"
---
# <a name="highresolutioniconurl-element"></a>HighResolutionIconUrl 元素

指定用于表示插入 UX 中的 Office 外接程序和高 DPI 屏幕上的 Office 应用商店的图像的 URL。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>可以包含

[Override](override.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**描述**|
|:-----|:-----|:-----|:-----|
|DefaultValue|字符串 (URL)|必需|指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|

## <a name="remarks"></a>注解

对于邮件外接程序，该图标将显示在 "**文件** > **管理外接程序**" UI 中。 For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.

图像必须采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。 对于内容和任务窗格应用程序，推荐的图像分辨率是 64 x 64 像素。 对于邮件应用程序，图像必须是 128 x 128 像素。 有关详细信息，请参阅[在 AppSource 和 Office 中创建有效的应用一览](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)中的_为你的应用创建一致的视觉标识_部分。

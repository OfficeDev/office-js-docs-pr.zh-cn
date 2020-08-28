---
title: 清单文件中的 IconUrl 元素
description: IconUrl 元素指定代表用户插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 27001f4109b2dcf93ac71d0a931bb6b4a2b38f2f
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292277"
---
# <a name="iconurl-element"></a>IconUrl 元素

指定用于表示插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>可以包含

[Override](override.md)

## <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|DefaultValue|字符串|必需|指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|

## <a name="remarks"></a>注解

对于邮件外接程序，该图标将显示在 "**文件**  >  **管理外接程序**" ui (Outlook) 或**设置**"  >  (outlook 网页版) 中**管理外接程序**ui。 For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. 对于所有外接程序类型，如果您将外接程序发布到 AppSource，则还会在 [AppSource](https://appsource.microsoft.com)中使用该图标。

图像必须采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。 对于内容和任务窗格应用程序，指定的图像必须是 32 x 32 像素。 对于邮件应用程序，推荐的图像分辨率是 64 x 64 像素。 此外，还应指定一个图标，以便与使用 [HighResolutionIconUrl](highresolutioniconurl.md) 元素在高 DPI 屏幕上运行的 Office 客户端应用程序一起使用。 有关详细信息，请参阅[在 AppSource 和 Office 中创建有效的应用一览](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)中的_为你的应用创建一致的视觉标识_部分。

`IconUrl`当前不支持在运行时更改元素的值。
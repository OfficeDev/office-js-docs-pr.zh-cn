---
title: 清单文件中的 IconUrl 元素
description: IconUrl 元素指定图像 URL，该图像表示Office UX 和 Office 应用商店中的外接程序。
ms.date: 03/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8f4026c9f97fc147b65532e370e18327bc07076c
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152674"
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

对于邮件外接程序，图标显示在"文件管理外接程序"UI  (Outlook) 或设置管理外接程序  >    >  UI (Outlook 网页版) 。 For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. 对于所有加载项类型，如果你将加载项发布到 [AppSource，](https://appsource.microsoft.com)图标也会在 AppSource 中使用。

图像必须采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。 对于内容和任务窗格应用程序，指定的图像必须是 32 x 32 像素。 对于邮件应用程序，图像分辨率必须为 64 x 64 像素。 还应指定一个图标，用于使用[highResolutionIconUrl](highresolutioniconurl.md) Office在高 DPI 屏幕上运行的客户端应用程序。 有关详细信息，请参阅 [在 AppSource 和 Office 中创建有效的应用一览](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)中的 _为你的应用创建一致的视觉标识_ 部分。

当前不支持 `IconUrl` 在运行时更改 元素的值。

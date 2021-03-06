---
title: 图像强制要求集
description: 支持跨 Excel、PowerPoint 和 Word 使用 Office 加载项的图像强制要求集。
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 52ce46a46580500f5a292bf898674d4798378319
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505526"
---
# <a name="image-coercion-requirement-sets"></a>图像强制要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 支持在 () `Office.CoercionType.Image` 写入数据时转换为图像 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 图像。 支持以下应用程序：

- Windows 版 Excel 2013 及更高版本
- Mac 版 Excel 2016 及更高版本
- iPad 版 Excel
- OneNote 网页版
- Windows 版 PowerPoint 2013 及更高版本
- Mac 版 PowerPoint 2016 及更高版本
- PowerPoint 网页版
- iPad 版 PowerPoint
- Windows 版 Word 2013 及更高版本
- Mac 版 Word 2016 及更高版本
- Word 网页版
- iPad 版 Word

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 支持在 () 写入数据时转换为 SVG `Office.CoercionType.XmlSvg` [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 格式。 支持以下应用程序：

- Windows 版 Excel (Microsoft 365 订阅) 
- Mac 版 Excel (Microsoft 365 订阅) 
- Windows 版 PowerPoint (连接到 Microsoft 365 订阅) 
- Mac 版 PowerPoint (Microsoft 365 订阅) 
- PowerPoint 网页版
- Windows 版 Word (Microsoft 365 订阅) 
- Mac 版 Word (Microsoft 365 订阅) 

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求集](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)

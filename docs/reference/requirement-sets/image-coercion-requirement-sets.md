---
title: 图像强制要求集
description: 支持跨 Office、Excel 和 Word PowerPoint外接程序的图像强制要求集。
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 35fed16003fe217e6f1f53d8c790cf78547308cf
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939303"
---
# <a name="image-coercion-requirement-sets"></a>图像强制要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 允许转换为图像 () `Office.CoercionType.Image` 方法写入数据时创建 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) 图像。 支持以下应用程序。

- Excel 2013 及更高版本Windows
- Excel 2016 Mac 及更高版本
- iPad 版 Excel
- OneNote 网页版
- PowerPoint 2013 年 10 月及Windows
- PowerPoint 2016 Mac 及更高版本
- PowerPoint 网页版
- iPad 版 PowerPoint
- Windows 版 Word 2013 及更高版本
- Mac 版 Word 2016 及更高版本
- Word 网页版
- iPad 版 Word

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 支持在使用 () 写入数据时转换为 SVG `Office.CoercionType.XmlSvg` [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) 格式。 支持以下应用程序。

- Excel连接到Windows (订阅Microsoft 365时) 
- Excel Mac (连接到 Microsoft 365 订阅) 
- PowerPoint连接到Windows (订阅Microsoft 365时) 
- PowerPoint Mac (连接到 Microsoft 365 订阅) 
- PowerPoint 网页版
- Word on Windows (连接到 Microsoft 365 订阅) 
- Mac 版 Word (连接到 Microsoft 365 订阅) 

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)

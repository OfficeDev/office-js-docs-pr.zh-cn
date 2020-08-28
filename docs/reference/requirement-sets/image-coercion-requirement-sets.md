---
title: 图像强制要求集
description: 支持跨 Excel、PowerPoint 和 Word 的 Office 外接程序对图像强制要求集的支持。
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 7140099757c6e4b5ad405723d5fed95fded6d919
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293546"
---
# <a name="image-coercion-requirement-sets"></a>图像强制要求集

要求集是指各组已命名的 API 成员。 Office 外接程序使用清单中指定的要求集或使用运行时检查来确定 Office 应用程序是否支持加载项所需的 Api。 有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

使用 ImageCoercion 1.1，可以 `Office.CoercionType.Image` 在使用方法写入数据时转换为) 的图像 ([`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 。 支持以下应用程序：

- Excel 2013 及更高版本的 Windows
- Excel 2016 及更高版本 Mac
- iPad 版 Excel
- OneNote 网页版
- PowerPoint 2013 及更高版本 Windows
- PowerPoint 2016 及更高版本 Mac
- PowerPoint 网页版
- iPad 版 PowerPoint
- Windows 版 Word 2013 及更高版本
- Mac 版 Word 2016 及更高版本
- Word 网页版
- iPad 版 Word

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 支持在 `Office.CoercionType.XmlSvg` 使用方法写入数据时 () 转换为 SVG 格式 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 。 支持以下应用程序：

- 连接到 Microsoft 365 订阅的 Windows (上的 Excel) 
- 连接到 Microsoft 365 订阅的 Mac 上的 Excel () 
- 连接到 Microsoft 365 订阅的 Windows (上的 PowerPoint) 
- 连接到 Microsoft 365 订阅的 Mac 版上的 PowerPoint () 
- PowerPoint 网页版
- 连接到 Microsoft 365 订阅的 Windows (上的 Word) 
- 连接到 Microsoft 365 订阅的 Mac 上的 Word () 
- Word 网页版

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)

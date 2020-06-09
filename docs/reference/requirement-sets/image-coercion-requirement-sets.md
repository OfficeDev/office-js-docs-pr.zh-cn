---
title: 图像强制要求集
description: 支持跨 Excel、PowerPoint 和 Word 的 Office 外接程序对图像强制要求集的支持。
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: f2baf8115d6a43c6b713e9acfeb5928f8549c583
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611356"
---
# <a name="image-coercion-requirement-sets"></a>图像强制要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

`Office.CoercionType.Image`在使用方法写入数据时，ImageCoercion 1.1 支持转换为 image （） [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 。 支持以下主机：

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

ImageCoercion 1.2 支持 `Office.CoercionType.XmlSvg` 在使用方法写入数据时转换为 SVG 格式（） [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 。 支持以下主机：

- Windows 上的 Excel （连接到 Office 365 订阅）
- Mac 上的 Excel （连接到 Office 365 订阅）
- Windows 上的 PowerPoint （连接到 Office 365 订阅）
- PowerPoint on Mac （连接到 Office 365 订阅）
- PowerPoint 网页版
- Windows 上的 Word （连接到 Office 365 订阅）
- Mac 上的 Word （连接到 Office 365 订阅）
- Word 网页版

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 主机和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)

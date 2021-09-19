---
title: 图像强制要求集
description: 支持跨 Office、PowerPoint 和 Word 的外接程序Excel强制要求集。
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 0f0b80c0af8213eaa9e3695373ddc037c2e60cc3
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/19/2021
ms.locfileid: "59448710"
---
# <a name="image-coercion-requirement-sets"></a>图像强制要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 允许转换为 () `Office.CoercionType.Image` 写入数据时显示 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) 的图像图像。 支持以下应用程序。

- Excel 2013 及更高版本Windows
- Excel 2016 Mac 及更高版本
- iPad 版 Excel
- OneNote 网页版
- PowerPoint 2013 及更高版本Windows
- PowerPoint 2016 Mac 及更高版本
- PowerPoint 网页版
- iPad 版 PowerPoint
- Windows 版 Word 2013 及更高版本
- Mac 版 Word 2016 及更高版本
- Word 网页版
- iPad 版 Word

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 支持在使用 () 写入数据时转换为 SVG `Office.CoercionType.XmlSvg` [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) 格式。 支持以下应用程序。

- Excel 2021 年 10 月及Windows
- Excel Mac 上的 2021 年及更高版本
- PowerPoint 2021 年 10 月及Windows
- PowerPoint Mac 上的 2021 年及更高版本
- PowerPoint 网页版
- Word 2021 及更高版本Windows
- Mac 上的 Word 2021 及更高版本

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)

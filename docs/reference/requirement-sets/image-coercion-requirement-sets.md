---
title: 图像强制要求集
description: 支持跨 Excel、PowerPoint 和 Word 的 Office 外接程序对图像强制要求集的支持。
ms.date: 07/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: bffe6c074d9e0734299d0087f2488524875931ed
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940842"
---
# <a name="image-coercion-requirement-sets"></a>图像强制要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Office 外接程序在多个 Office 版本中运行。 下表列出了图像强制要求集、支持该要求集的 Office 主机应用程序, 以及 Office 应用程序的内部版本号或版本号。

## <a name="imagecoercion-11"></a>ImageCoercion 1。1

在使用[`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)方法写入数据时, ImageCoercion`Office.CoercionType.Image`1.1 支持转换为 image ()。 支持以下主机:

- Excel 2013 及更高版本的 Windows
- Excel 2016 及更高版本 Mac
- 在 web 上的 Excel
- IPad 上的 Excel
- 在 web 上的 OneNote
- PowerPoint 2013 及更高版本 Windows
- PowerPoint 2016 及更高版本 Mac
- 在 web 上的 PowerPoint
- IPad 上的 PowerPoint
- Word 2013 及更高版本的 Windows
- Word 2016 及更高版本 Mac
- 在 web 上的 Word
- iPad 上的 Word

## <a name="imagecoercion-12"></a>ImageCoercion 1。2

ImageCoercion 1.2 支持在使用`Office.CoercionType.XmlSvg` [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)方法写入数据时转换为 SVG 格式 ()。 支持以下主机:

- Windows 上的 Excel (连接到 Office 365 订阅)
- Mac 上的 Excel (连接到 Office 365 订阅)
- 在 web 上的 Excel
- Windows 上的 PowerPoint (连接到 Office 365 订阅)
- PowerPoint on Mac (连接到 Office 365 订阅)
- 在 web 上的 PowerPoint
- Windows 上的 Word (连接到 Office 365 订阅)
- Mac 上的 Word (连接到 Office 365 订阅)
- 在 web 上的 Word

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](/office/dev/add-ins/develop/add-in-manifests)

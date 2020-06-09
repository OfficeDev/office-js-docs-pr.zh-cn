---
title: 清单文件中的 SupportUrl 元素
description: SupportUrl 元素指定为外接程序提供支持信息的页面的 URL。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f75ee811699823a501ac594e66daaaf3f93c2782
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608703"
---
# <a name="supporturl-element"></a>SupportUrl 元素

指定提供外接程序支持信息的页面的 URL。

## <a name="syntax"></a>语法

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

|  元素 | 必需 | Description  |
|:-----|:-----|:-----|
|  [Override](override.md)   | 否 | 指定其他区域设置 URL 的设置 |

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**描述**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必需|指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|

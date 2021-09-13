---
title: 清单文件中的 SupportUrl 元素
description: SupportUrl 元素指定为您的外接程序提供支持信息的页面的 URL。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 2ea515aa61ed5bf9e22d6316a76fa4b5e51493f3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149442"
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

|  元素 | 必需 | 说明  |
|:-----|:-----|:-----|
|  [Override](override.md)   | 否 | 指定其他区域设置 URL 的设置 |

## <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必需|指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|

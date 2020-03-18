---
title: 清单文件中的 Override 元素
description: Override 元素使您能够为其他区域设置指定设置的值。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 39e706dc981d405fcfcc508626578f34931efbcb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718025"
---
# <a name="override-element"></a>Override 元素

提供一种为其他区域设置指定某设置的值的方法。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>包含于

|**Element**|
|:-----|
|[CitationText](citationtext.md)|
|[说明](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**描述**|
|:-----|:-----|:-----|:-----|
|区域设置|string|必需|为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。|
|值|字符串|必需|指定表示为指定区域设置的设置的值。|

## <a name="see-also"></a>另请参阅

- [Office 外接程序的本地化](../../develop/localization.md)

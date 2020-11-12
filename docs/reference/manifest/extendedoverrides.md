---
title: 清单文件中的 ExtendedOverrides 元素
description: 指定清单的 JSON 格式扩展名的 Url。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 76491af34d1caf0ec266826df97a5363e336b85d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996676"
---
# <a name="extendedoverrides-element"></a>ExtendedOverrides 元素

指定用于扩展清单的 JSON 格式文件的完整 Url。

**外接程序类型：** 任务窗格

## <a name="syntax"></a>语法

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

|元素|内容|邮件|任务窗格|
|:-----|:-----|:-----|:-----|
|[等级](tokens.md)|||x|

## <a name="attributes"></a>属性

|属性|说明|
|:-----|:-----|
|Url (必需的) | 扩展替代 JSON 文件的完整 URL。 这可以是使用 [令牌](tokens.md) 元素所定义的令牌的 URL 模板。|
|ResourcesUrl (可选)  | 为属性中指定的文件提供补充资源（如本地化字符串）的文件的完整 URL `Url` 。 这可以是使用 [令牌](tokens.md) 元素所定义的令牌的 URL 模板。|

## <a name="example"></a>示例

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```

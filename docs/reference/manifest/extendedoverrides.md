---
title: 清单文件中 ExtendedOverrides 元素
description: 指定清单的 JSON 格式扩展的 URL。
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: f2b9ea409763119b5bec5286ecdc5f15c94c49e6312a13209197e6457353f369
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57083580"
---
# <a name="extendedoverrides-element"></a>ExtendedOverrides 元素

指定用于扩展清单的 JSON 格式文件的完整 URL。 有关使用此元素及其后代元素的详细信息，请参阅使用清单 [的扩展替代](../../develop/extended-overrides.md)。

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
|[令牌](tokens.md)|||x|

## <a name="attributes"></a>属性

|属性|说明|
|:-----|:-----|
|Url (必需) | 扩展的完整 URL 重写 JSON 文件。 将来，此值可能是使用 [Tokens](tokens.md) 元素定义的令牌的 URL 模板。 请参阅 [示例](#examples)。|
|ResourcesUrl (可选)  | 为 属性中指定的文件提供补充资源（如本地化字符串）的文件的完整 `Url` URL。 这可能是使用 [Tokens](tokens.md) 元素定义的令牌的 URL 模板。|

## <a name="examples"></a>示例

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

将来，此值可能是使用 [Tokens](tokens.md) 元素定义的令牌的 URL 模板。 示例如下。

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

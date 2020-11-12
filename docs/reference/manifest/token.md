---
title: 清单文件中的 Token 元素
description: 指定可与清单中的 URL 模板一起使用的令牌或通配符。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5e26af44c566ab09ac81c8194e1ae7d85aaac327
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996675"
---
# <a name="token-element"></a>Token 元素

定义单个 URL 标记。

**外接程序类型：** 任务窗格

## <a name="syntax"></a>语法

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a>包含于

[等级](tokens.md)

## <a name="can-contain"></a>可以包含

|元素|内容|邮件|任务窗格|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## <a name="attributes"></a>属性

|属性|说明|
|:-----|:-----|
|DefaultValue|此令牌的默认值（如果任何子元素中没有匹配的条件） `<Override>` 。|
|名称|令牌名称。 此名称是用户定义的。 令牌的类型由 type 属性决定。|
|xsi:type|定义令牌的种类。 此属性应设置为以下其中一个：  `"RequirementsToken"` 、或  `"LocaleToken"` 。|

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
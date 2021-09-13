---
title: 清单文件中标记元素
description: 指定可用于清单中的 URL 模板的令牌或通配符。
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: 69f626f5f6f57dd155756812bcd56267a1da3ffa
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149430"
---
# <a name="token-element"></a>Token 元素

定义单个 URL 令牌。 有关此元素的使用详细信息，请参阅使用 [清单的扩展替代](../../develop/extended-overrides.md)。

**外接程序类型：** 任务窗格

## <a name="syntax"></a>语法

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a>包含于

[令牌](tokens.md)

## <a name="can-contain"></a>可以包含

|元素|内容|邮件|任务窗格|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## <a name="attributes"></a>属性

|属性|说明|
|:-----|:-----|
|DefaultValue|如果任何子元素中的条件都匹配，则此令牌 `<Override>` 的默认值。|
|名称|令牌名称。 此名称是用户定义的。 令牌的类型由 type 属性确定。|
|xsi:type|定义令牌类型。 此属性应设置为：或 `"RequirementsToken"` 。 `"LocaleToken"`|

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
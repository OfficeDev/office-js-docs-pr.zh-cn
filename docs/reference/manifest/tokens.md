---
title: 清单文件中 Tokens 元素
description: 指定可用于清单中的 URL 模板的标记或通配符。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 8680b985068c44e93f601a2b24e2f28899eb483d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505323"
---
# <a name="tokens-element"></a>Tokens 元素

定义可以在模板 URL 中使用的令牌。 有关使用此元素的信息，请参阅使用清单 [的扩展重写](../../develop/extended-overrides.md)。

**外接程序类型：** 任务窗格

## <a name="syntax"></a>语法

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a>包含于

[ExtendedOverrides](extendedoverrides.md)

## <a name="must-contain"></a>必须包含

|元素|内容|邮件|任务窗格|
|:-----|:-----|:-----|:-----|
|[标记](token.md)|||x|

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
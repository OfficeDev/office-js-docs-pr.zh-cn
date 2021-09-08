---
title: 清单文件的 Tokens 元素
description: 指定可用于清单中的 URL 模板的标记或通配符。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 8680b985068c44e93f601a2b24e2f28899eb483d
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938230"
---
# <a name="tokens-element"></a>Tokens 元素

定义可以在模板 URL 中使用的令牌。 有关此元素的使用详细信息，请参阅使用 [清单的扩展替代](../../develop/extended-overrides.md)。

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
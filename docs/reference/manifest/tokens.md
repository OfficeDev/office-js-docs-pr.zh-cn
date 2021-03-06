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
# <a name="tokens-element"></a><span data-ttu-id="0b694-103">Tokens 元素</span><span class="sxs-lookup"><span data-stu-id="0b694-103">Tokens element</span></span>

<span data-ttu-id="0b694-104">定义可以在模板 URL 中使用的令牌。</span><span class="sxs-lookup"><span data-stu-id="0b694-104">Defines tokens that could be used in template URLs.</span></span> <span data-ttu-id="0b694-105">有关使用此元素的信息，请参阅使用清单 [的扩展重写](../../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="0b694-105">For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="0b694-106">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b694-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="0b694-107">语法</span><span class="sxs-lookup"><span data-stu-id="0b694-107">Syntax</span></span>

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a><span data-ttu-id="0b694-108">包含于</span><span class="sxs-lookup"><span data-stu-id="0b694-108">Contained in</span></span>

[<span data-ttu-id="0b694-109">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="0b694-109">ExtendedOverrides</span></span>](extendedoverrides.md)

## <a name="must-contain"></a><span data-ttu-id="0b694-110">必须包含</span><span class="sxs-lookup"><span data-stu-id="0b694-110">Must contain</span></span>

|<span data-ttu-id="0b694-111">元素</span><span class="sxs-lookup"><span data-stu-id="0b694-111">Element</span></span>|<span data-ttu-id="0b694-112">内容</span><span class="sxs-lookup"><span data-stu-id="0b694-112">Content</span></span>|<span data-ttu-id="0b694-113">邮件</span><span class="sxs-lookup"><span data-stu-id="0b694-113">Mail</span></span>|<span data-ttu-id="0b694-114">任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b694-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="0b694-115">标记</span><span class="sxs-lookup"><span data-stu-id="0b694-115">Token</span></span>](token.md)|||<span data-ttu-id="0b694-116">x</span><span class="sxs-lookup"><span data-stu-id="0b694-116">x</span></span>|

## <a name="example"></a><span data-ttu-id="0b694-117">示例</span><span class="sxs-lookup"><span data-stu-id="0b694-117">Example</span></span>

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
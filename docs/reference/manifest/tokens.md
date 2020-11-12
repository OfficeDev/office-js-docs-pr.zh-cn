---
title: 清单文件中的标记元素
description: 指定可与清单中的 URL 模板一起使用的标记或通配符。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: a50de7c2c3e8ebeb9425c1677a94bbcc62281d3b
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996673"
---
# <a name="tokens-element"></a><span data-ttu-id="6d19c-103">标记元素</span><span class="sxs-lookup"><span data-stu-id="6d19c-103">Tokens element</span></span>

<span data-ttu-id="6d19c-104">定义可在模板 Url 中使用的标记。</span><span class="sxs-lookup"><span data-stu-id="6d19c-104">Defines tokens that could be used in template URLs.</span></span>

<span data-ttu-id="6d19c-105">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="6d19c-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="6d19c-106">语法</span><span class="sxs-lookup"><span data-stu-id="6d19c-106">Syntax</span></span>

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a><span data-ttu-id="6d19c-107">包含于</span><span class="sxs-lookup"><span data-stu-id="6d19c-107">Contained in</span></span>

[<span data-ttu-id="6d19c-108">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="6d19c-108">ExtendedOverrides</span></span>](extendedoverrides.md)

## <a name="must-contain"></a><span data-ttu-id="6d19c-109">必须包含</span><span class="sxs-lookup"><span data-stu-id="6d19c-109">Must contain</span></span>

|<span data-ttu-id="6d19c-110">元素</span><span class="sxs-lookup"><span data-stu-id="6d19c-110">Element</span></span>|<span data-ttu-id="6d19c-111">内容</span><span class="sxs-lookup"><span data-stu-id="6d19c-111">Content</span></span>|<span data-ttu-id="6d19c-112">邮件</span><span class="sxs-lookup"><span data-stu-id="6d19c-112">Mail</span></span>|<span data-ttu-id="6d19c-113">任务窗格</span><span class="sxs-lookup"><span data-stu-id="6d19c-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="6d19c-114">标记</span><span class="sxs-lookup"><span data-stu-id="6d19c-114">Token</span></span>](token.md)|||<span data-ttu-id="6d19c-115">x</span><span class="sxs-lookup"><span data-stu-id="6d19c-115">x</span></span>|

## <a name="example"></a><span data-ttu-id="6d19c-116">示例</span><span class="sxs-lookup"><span data-stu-id="6d19c-116">Example</span></span>

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
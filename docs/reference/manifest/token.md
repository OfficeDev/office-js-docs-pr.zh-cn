---
title: 清单文件中标记元素
description: 指定可用于清单中的 URL 模板的令牌或通配符。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 48078f8211a8fd3f0e3f9d7c3f3aabd1d31b0a6d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505365"
---
# <a name="token-element"></a><span data-ttu-id="7d281-103">Token 元素</span><span class="sxs-lookup"><span data-stu-id="7d281-103">Token element</span></span>

<span data-ttu-id="7d281-104">定义单个 URL 令牌。</span><span class="sxs-lookup"><span data-stu-id="7d281-104">Defines an individual URL token.</span></span> <span data-ttu-id="7d281-105">有关使用此元素的信息，请参阅使用清单 [的扩展重写](../../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="7d281-105">For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="7d281-106">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7d281-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="7d281-107">语法</span><span class="sxs-lookup"><span data-stu-id="7d281-107">Syntax</span></span>

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a><span data-ttu-id="7d281-108">包含于</span><span class="sxs-lookup"><span data-stu-id="7d281-108">Contained in</span></span>

[<span data-ttu-id="7d281-109">令牌</span><span class="sxs-lookup"><span data-stu-id="7d281-109">Tokens</span></span>](tokens.md)

## <a name="can-contain"></a><span data-ttu-id="7d281-110">可以包含</span><span class="sxs-lookup"><span data-stu-id="7d281-110">Can contain</span></span>

|<span data-ttu-id="7d281-111">元素</span><span class="sxs-lookup"><span data-stu-id="7d281-111">Element</span></span>|<span data-ttu-id="7d281-112">内容</span><span class="sxs-lookup"><span data-stu-id="7d281-112">Content</span></span>|<span data-ttu-id="7d281-113">邮件</span><span class="sxs-lookup"><span data-stu-id="7d281-113">Mail</span></span>|<span data-ttu-id="7d281-114">任务窗格</span><span class="sxs-lookup"><span data-stu-id="7d281-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="7d281-115">Override</span><span class="sxs-lookup"><span data-stu-id="7d281-115">Override</span></span>](override.md)|||<span data-ttu-id="7d281-116">x</span><span class="sxs-lookup"><span data-stu-id="7d281-116">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="7d281-117">属性</span><span class="sxs-lookup"><span data-stu-id="7d281-117">Attributes</span></span>

|<span data-ttu-id="7d281-118">属性</span><span class="sxs-lookup"><span data-stu-id="7d281-118">Attribute</span></span>|<span data-ttu-id="7d281-119">说明</span><span class="sxs-lookup"><span data-stu-id="7d281-119">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="7d281-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="7d281-120">DefaultValue</span></span>|<span data-ttu-id="7d281-121">如果任何子元素中没有任何条件匹配，则此令牌 `<Override>` 的默认值。</span><span class="sxs-lookup"><span data-stu-id="7d281-121">Default value for this token if no condition in any child `<Override>` element matches.</span></span>|
|<span data-ttu-id="7d281-122">名称</span><span class="sxs-lookup"><span data-stu-id="7d281-122">Name</span></span>|<span data-ttu-id="7d281-123">令牌名称。</span><span class="sxs-lookup"><span data-stu-id="7d281-123">Token name.</span></span> <span data-ttu-id="7d281-124">此名称是用户定义的。</span><span class="sxs-lookup"><span data-stu-id="7d281-124">This name is user-defined.</span></span> <span data-ttu-id="7d281-125">令牌的类型由类型属性决定。</span><span class="sxs-lookup"><span data-stu-id="7d281-125">The type of the token is determined by the type attribute.</span></span>|
|<span data-ttu-id="7d281-126">xsi:type</span><span class="sxs-lookup"><span data-stu-id="7d281-126">xsi:type</span></span>|<span data-ttu-id="7d281-127">定义令牌类型。</span><span class="sxs-lookup"><span data-stu-id="7d281-127">Defines the kind of Token.</span></span> <span data-ttu-id="7d281-128">此属性应设置为： 或 `"RequirementsToken"` 。 `"LocaleToken"`</span><span class="sxs-lookup"><span data-stu-id="7d281-128">This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.</span></span>|

## <a name="example"></a><span data-ttu-id="7d281-129">示例</span><span class="sxs-lookup"><span data-stu-id="7d281-129">Example</span></span>

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
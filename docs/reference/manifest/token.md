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
# <a name="token-element"></a><span data-ttu-id="919d4-103">Token 元素</span><span class="sxs-lookup"><span data-stu-id="919d4-103">Token element</span></span>

<span data-ttu-id="919d4-104">定义单个 URL 标记。</span><span class="sxs-lookup"><span data-stu-id="919d4-104">Defines an individual URL token.</span></span>

<span data-ttu-id="919d4-105">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="919d4-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="919d4-106">语法</span><span class="sxs-lookup"><span data-stu-id="919d4-106">Syntax</span></span>

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a><span data-ttu-id="919d4-107">包含于</span><span class="sxs-lookup"><span data-stu-id="919d4-107">Contained in</span></span>

[<span data-ttu-id="919d4-108">等级</span><span class="sxs-lookup"><span data-stu-id="919d4-108">Tokens</span></span>](tokens.md)

## <a name="can-contain"></a><span data-ttu-id="919d4-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="919d4-109">Can contain</span></span>

|<span data-ttu-id="919d4-110">元素</span><span class="sxs-lookup"><span data-stu-id="919d4-110">Element</span></span>|<span data-ttu-id="919d4-111">内容</span><span class="sxs-lookup"><span data-stu-id="919d4-111">Content</span></span>|<span data-ttu-id="919d4-112">邮件</span><span class="sxs-lookup"><span data-stu-id="919d4-112">Mail</span></span>|<span data-ttu-id="919d4-113">任务窗格</span><span class="sxs-lookup"><span data-stu-id="919d4-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="919d4-114">Override</span><span class="sxs-lookup"><span data-stu-id="919d4-114">Override</span></span>](override.md)|||<span data-ttu-id="919d4-115">x</span><span class="sxs-lookup"><span data-stu-id="919d4-115">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="919d4-116">属性</span><span class="sxs-lookup"><span data-stu-id="919d4-116">Attributes</span></span>

|<span data-ttu-id="919d4-117">属性</span><span class="sxs-lookup"><span data-stu-id="919d4-117">Attribute</span></span>|<span data-ttu-id="919d4-118">说明</span><span class="sxs-lookup"><span data-stu-id="919d4-118">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="919d4-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="919d4-119">DefaultValue</span></span>|<span data-ttu-id="919d4-120">此令牌的默认值（如果任何子元素中没有匹配的条件） `<Override>` 。</span><span class="sxs-lookup"><span data-stu-id="919d4-120">Default value for this token if no condition in any child `<Override>` element matches.</span></span>|
|<span data-ttu-id="919d4-121">名称</span><span class="sxs-lookup"><span data-stu-id="919d4-121">Name</span></span>|<span data-ttu-id="919d4-122">令牌名称。</span><span class="sxs-lookup"><span data-stu-id="919d4-122">Token name.</span></span> <span data-ttu-id="919d4-123">此名称是用户定义的。</span><span class="sxs-lookup"><span data-stu-id="919d4-123">This name is user-defined.</span></span> <span data-ttu-id="919d4-124">令牌的类型由 type 属性决定。</span><span class="sxs-lookup"><span data-stu-id="919d4-124">The type of the token is determined by the type attribute.</span></span>|
|<span data-ttu-id="919d4-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="919d4-125">xsi:type</span></span>|<span data-ttu-id="919d4-126">定义令牌的种类。</span><span class="sxs-lookup"><span data-stu-id="919d4-126">Defines the kind of Token.</span></span> <span data-ttu-id="919d4-127">此属性应设置为以下其中一个：  `"RequirementsToken"` 、或  `"LocaleToken"` 。</span><span class="sxs-lookup"><span data-stu-id="919d4-127">This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.</span></span>|

## <a name="example"></a><span data-ttu-id="919d4-128">示例</span><span class="sxs-lookup"><span data-stu-id="919d4-128">Example</span></span>

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
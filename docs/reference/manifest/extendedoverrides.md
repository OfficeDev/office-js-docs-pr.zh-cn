---
title: 清单文件中 ExtendedOverrides 元素
description: 指定清单的 JSON 格式扩展的 URL。
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: f433c9c5604f3fae35580ba20780ea6fe91401c7
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505470"
---
# <a name="extendedoverrides-element"></a><span data-ttu-id="f0893-103">ExtendedOverrides 元素</span><span class="sxs-lookup"><span data-stu-id="f0893-103">ExtendedOverrides element</span></span>

<span data-ttu-id="f0893-104">指定扩展清单的 JSON 格式文件的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="f0893-104">Specifies the full URLs for JSON-formatted files that extend the manifest.</span></span> <span data-ttu-id="f0893-105">有关使用此元素及其后代元素的详细信息，请参阅使用清单 [的扩展重写](../../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="f0893-105">For detailed information about the use of this element and its descendent elements, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="f0893-106">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f0893-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="f0893-107">语法</span><span class="sxs-lookup"><span data-stu-id="f0893-107">Syntax</span></span>

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a><span data-ttu-id="f0893-108">包含于</span><span class="sxs-lookup"><span data-stu-id="f0893-108">Contained in</span></span>

[<span data-ttu-id="f0893-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f0893-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="f0893-110">可以包含</span><span class="sxs-lookup"><span data-stu-id="f0893-110">Can contain</span></span>

|<span data-ttu-id="f0893-111">元素</span><span class="sxs-lookup"><span data-stu-id="f0893-111">Element</span></span>|<span data-ttu-id="f0893-112">内容</span><span class="sxs-lookup"><span data-stu-id="f0893-112">Content</span></span>|<span data-ttu-id="f0893-113">邮件</span><span class="sxs-lookup"><span data-stu-id="f0893-113">Mail</span></span>|<span data-ttu-id="f0893-114">任务窗格</span><span class="sxs-lookup"><span data-stu-id="f0893-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="f0893-115">令牌</span><span class="sxs-lookup"><span data-stu-id="f0893-115">Tokens</span></span>](tokens.md)|||<span data-ttu-id="f0893-116">x</span><span class="sxs-lookup"><span data-stu-id="f0893-116">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="f0893-117">属性</span><span class="sxs-lookup"><span data-stu-id="f0893-117">Attributes</span></span>

|<span data-ttu-id="f0893-118">属性</span><span class="sxs-lookup"><span data-stu-id="f0893-118">Attribute</span></span>|<span data-ttu-id="f0893-119">说明</span><span class="sxs-lookup"><span data-stu-id="f0893-119">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="f0893-120">Url (必需) </span><span class="sxs-lookup"><span data-stu-id="f0893-120">Url (required)</span></span>| <span data-ttu-id="f0893-121">扩展的 URL 将覆盖 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="f0893-121">The full URL of the extended overrides JSON file.</span></span> <span data-ttu-id="f0893-122">将来，此值可能是使用 [Tokens](tokens.md) 元素定义的令牌的 URL 模板。</span><span class="sxs-lookup"><span data-stu-id="f0893-122">In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span> <span data-ttu-id="f0893-123">请参阅 [示例](#examples)。</span><span class="sxs-lookup"><span data-stu-id="f0893-123">See [Examples](#examples).</span></span>|
|<span data-ttu-id="f0893-124">ResourcesUrl (可选) </span><span class="sxs-lookup"><span data-stu-id="f0893-124">ResourcesUrl (optional)</span></span> | <span data-ttu-id="f0893-125">为属性中指定的文件提供补充资源（如本地化字符串）的文件的完整 `Url` URL。</span><span class="sxs-lookup"><span data-stu-id="f0893-125">The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute.</span></span> <span data-ttu-id="f0893-126">这可能是使用 [Tokens](tokens.md) 元素定义的令牌的 URL 模板。</span><span class="sxs-lookup"><span data-stu-id="f0893-126">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|

## <a name="examples"></a><span data-ttu-id="f0893-127">示例</span><span class="sxs-lookup"><span data-stu-id="f0893-127">Examples</span></span>

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="f0893-128">将来，此值可能是使用 [Tokens](tokens.md) 元素定义的令牌的 URL 模板。</span><span class="sxs-lookup"><span data-stu-id="f0893-128">In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span> <span data-ttu-id="f0893-129">示例如下。</span><span class="sxs-lookup"><span data-stu-id="f0893-129">The following is an example.</span></span>

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

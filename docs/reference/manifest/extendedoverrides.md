---
title: 清单文件中的 ExtendedOverrides 元素
description: 指定清单的 JSON 格式扩展名的 Url。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 76491af34d1caf0ec266826df97a5363e336b85d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996676"
---
# <a name="extendedoverrides-element"></a><span data-ttu-id="e0c47-103">ExtendedOverrides 元素</span><span class="sxs-lookup"><span data-stu-id="e0c47-103">ExtendedOverrides element</span></span>

<span data-ttu-id="e0c47-104">指定用于扩展清单的 JSON 格式文件的完整 Url。</span><span class="sxs-lookup"><span data-stu-id="e0c47-104">Specifies the full URLs for JSON-formatted files that extend the manifest.</span></span>

<span data-ttu-id="e0c47-105">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e0c47-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="e0c47-106">语法</span><span class="sxs-lookup"><span data-stu-id="e0c47-106">Syntax</span></span>

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a><span data-ttu-id="e0c47-107">包含于</span><span class="sxs-lookup"><span data-stu-id="e0c47-107">Contained in</span></span>

[<span data-ttu-id="e0c47-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="e0c47-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="e0c47-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="e0c47-109">Can contain</span></span>

|<span data-ttu-id="e0c47-110">元素</span><span class="sxs-lookup"><span data-stu-id="e0c47-110">Element</span></span>|<span data-ttu-id="e0c47-111">内容</span><span class="sxs-lookup"><span data-stu-id="e0c47-111">Content</span></span>|<span data-ttu-id="e0c47-112">邮件</span><span class="sxs-lookup"><span data-stu-id="e0c47-112">Mail</span></span>|<span data-ttu-id="e0c47-113">任务窗格</span><span class="sxs-lookup"><span data-stu-id="e0c47-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="e0c47-114">等级</span><span class="sxs-lookup"><span data-stu-id="e0c47-114">Tokens</span></span>](tokens.md)|||<span data-ttu-id="e0c47-115">x</span><span class="sxs-lookup"><span data-stu-id="e0c47-115">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="e0c47-116">属性</span><span class="sxs-lookup"><span data-stu-id="e0c47-116">Attributes</span></span>

|<span data-ttu-id="e0c47-117">属性</span><span class="sxs-lookup"><span data-stu-id="e0c47-117">Attribute</span></span>|<span data-ttu-id="e0c47-118">说明</span><span class="sxs-lookup"><span data-stu-id="e0c47-118">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="e0c47-119">Url (必需的) </span><span class="sxs-lookup"><span data-stu-id="e0c47-119">Url (required)</span></span>| <span data-ttu-id="e0c47-120">扩展替代 JSON 文件的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="e0c47-120">The full URL of the extended overrides JSON file.</span></span> <span data-ttu-id="e0c47-121">这可以是使用 [令牌](tokens.md) 元素所定义的令牌的 URL 模板。</span><span class="sxs-lookup"><span data-stu-id="e0c47-121">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|
|<span data-ttu-id="e0c47-122">ResourcesUrl (可选) </span><span class="sxs-lookup"><span data-stu-id="e0c47-122">ResourcesUrl (optional)</span></span> | <span data-ttu-id="e0c47-123">为属性中指定的文件提供补充资源（如本地化字符串）的文件的完整 URL `Url` 。</span><span class="sxs-lookup"><span data-stu-id="e0c47-123">The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute.</span></span> <span data-ttu-id="e0c47-124">这可以是使用 [令牌](tokens.md) 元素所定义的令牌的 URL 模板。</span><span class="sxs-lookup"><span data-stu-id="e0c47-124">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|

## <a name="example"></a><span data-ttu-id="e0c47-125">示例</span><span class="sxs-lookup"><span data-stu-id="e0c47-125">Example</span></span>

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

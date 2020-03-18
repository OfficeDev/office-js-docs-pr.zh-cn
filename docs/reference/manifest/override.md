---
title: 清单文件中的 Override 元素
description: Override 元素使您能够为其他区域设置指定设置的值。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 39e706dc981d405fcfcc508626578f34931efbcb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718025"
---
# <a name="override-element"></a><span data-ttu-id="9de50-103">Override 元素</span><span class="sxs-lookup"><span data-stu-id="9de50-103">Override element</span></span>

<span data-ttu-id="9de50-104">提供一种为其他区域设置指定某设置的值的方法。</span><span class="sxs-lookup"><span data-stu-id="9de50-104">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="9de50-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="9de50-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9de50-106">语法</span><span class="sxs-lookup"><span data-stu-id="9de50-106">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="9de50-107">包含于</span><span class="sxs-lookup"><span data-stu-id="9de50-107">Contained in</span></span>

|<span data-ttu-id="9de50-108">**Element**</span><span class="sxs-lookup"><span data-stu-id="9de50-108">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="9de50-109">CitationText</span><span class="sxs-lookup"><span data-stu-id="9de50-109">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="9de50-110">说明</span><span class="sxs-lookup"><span data-stu-id="9de50-110">Description</span></span>](description.md)|
|[<span data-ttu-id="9de50-111">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="9de50-111">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="9de50-112">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="9de50-112">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="9de50-113">DisplayName</span><span class="sxs-lookup"><span data-stu-id="9de50-113">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="9de50-114">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="9de50-114">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="9de50-115">IconUrl</span><span class="sxs-lookup"><span data-stu-id="9de50-115">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="9de50-116">QueryUri</span><span class="sxs-lookup"><span data-stu-id="9de50-116">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="9de50-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9de50-117">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="9de50-118">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="9de50-118">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="9de50-119">属性</span><span class="sxs-lookup"><span data-stu-id="9de50-119">Attributes</span></span>

|<span data-ttu-id="9de50-120">**属性**</span><span class="sxs-lookup"><span data-stu-id="9de50-120">**Attribute**</span></span>|<span data-ttu-id="9de50-121">**类型**</span><span class="sxs-lookup"><span data-stu-id="9de50-121">**Type**</span></span>|<span data-ttu-id="9de50-122">**必需**</span><span class="sxs-lookup"><span data-stu-id="9de50-122">**Required**</span></span>|<span data-ttu-id="9de50-123">**描述**</span><span class="sxs-lookup"><span data-stu-id="9de50-123">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="9de50-124">区域设置</span><span class="sxs-lookup"><span data-stu-id="9de50-124">Locale</span></span>|<span data-ttu-id="9de50-125">string</span><span class="sxs-lookup"><span data-stu-id="9de50-125">string</span></span>|<span data-ttu-id="9de50-126">必需</span><span class="sxs-lookup"><span data-stu-id="9de50-126">required</span></span>|<span data-ttu-id="9de50-127">为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。</span><span class="sxs-lookup"><span data-stu-id="9de50-127">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="9de50-128">值</span><span class="sxs-lookup"><span data-stu-id="9de50-128">Value</span></span>|<span data-ttu-id="9de50-129">字符串</span><span class="sxs-lookup"><span data-stu-id="9de50-129">string</span></span>|<span data-ttu-id="9de50-130">必需</span><span class="sxs-lookup"><span data-stu-id="9de50-130">required</span></span>|<span data-ttu-id="9de50-131">指定表示为指定区域设置的设置的值。</span><span class="sxs-lookup"><span data-stu-id="9de50-131">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="9de50-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9de50-132">See also</span></span>

- [<span data-ttu-id="9de50-133">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="9de50-133">Localization for Office Add-ins</span></span>](../../develop/localization.md)

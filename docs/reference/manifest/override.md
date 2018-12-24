---
title: 清单文件中的 Override 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1d2400312f12116b1ac5f4010135541e783dcc7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432863"
---
# <a name="override-element"></a><span data-ttu-id="aaae5-102">Override 元素</span><span class="sxs-lookup"><span data-stu-id="aaae5-102">Override element</span></span>

<span data-ttu-id="aaae5-103">提供一种为其他区域设置指定某设置的值的方法。</span><span class="sxs-lookup"><span data-stu-id="aaae5-103">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="aaae5-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="aaae5-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="aaae5-105">语法</span><span class="sxs-lookup"><span data-stu-id="aaae5-105">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="aaae5-106">包含于</span><span class="sxs-lookup"><span data-stu-id="aaae5-106">Contained in</span></span>

|<span data-ttu-id="aaae5-107">**Element**</span><span class="sxs-lookup"><span data-stu-id="aaae5-107">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="aaae5-108">CitationText</span><span class="sxs-lookup"><span data-stu-id="aaae5-108">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="aaae5-109">说明</span><span class="sxs-lookup"><span data-stu-id="aaae5-109">Description</span></span>](description.md)|
|[<span data-ttu-id="aaae5-110">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="aaae5-110">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="aaae5-111">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="aaae5-111">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="aaae5-112">DisplayName</span><span class="sxs-lookup"><span data-stu-id="aaae5-112">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="aaae5-113">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="aaae5-113">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="aaae5-114">IconUrl</span><span class="sxs-lookup"><span data-stu-id="aaae5-114">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="aaae5-115">QueryUri</span><span class="sxs-lookup"><span data-stu-id="aaae5-115">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="aaae5-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="aaae5-116">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="aaae5-117">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="aaae5-117">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="aaae5-118">属性</span><span class="sxs-lookup"><span data-stu-id="aaae5-118">Attributes</span></span>

|<span data-ttu-id="aaae5-119">**属性**</span><span class="sxs-lookup"><span data-stu-id="aaae5-119">**Attribute**</span></span>|<span data-ttu-id="aaae5-120">**类型**</span><span class="sxs-lookup"><span data-stu-id="aaae5-120">**Type**</span></span>|<span data-ttu-id="aaae5-121">**必需**</span><span class="sxs-lookup"><span data-stu-id="aaae5-121">**Required**</span></span>|<span data-ttu-id="aaae5-122">**说明**</span><span class="sxs-lookup"><span data-stu-id="aaae5-122">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="aaae5-123">Locale</span><span class="sxs-lookup"><span data-stu-id="aaae5-123">Locale</span></span>|<span data-ttu-id="aaae5-124">字符串</span><span class="sxs-lookup"><span data-stu-id="aaae5-124">string</span></span>|<span data-ttu-id="aaae5-125">必需</span><span class="sxs-lookup"><span data-stu-id="aaae5-125">required</span></span>|<span data-ttu-id="aaae5-126">为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。</span><span class="sxs-lookup"><span data-stu-id="aaae5-126">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="aaae5-127">值</span><span class="sxs-lookup"><span data-stu-id="aaae5-127">Value</span></span>|<span data-ttu-id="aaae5-128">字符串</span><span class="sxs-lookup"><span data-stu-id="aaae5-128">string</span></span>|<span data-ttu-id="aaae5-129">必需</span><span class="sxs-lookup"><span data-stu-id="aaae5-129">required</span></span>|<span data-ttu-id="aaae5-130">指定表示为指定区域设置的设置的值。</span><span class="sxs-lookup"><span data-stu-id="aaae5-130">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="aaae5-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="aaae5-131">See also</span></span>

- [<span data-ttu-id="aaae5-132">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="aaae5-132">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    

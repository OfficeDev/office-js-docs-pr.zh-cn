---
title: 清单文件中的 Override 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 020ae490dacbb9b8c493dc022c23d0ebf311a1b9
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450448"
---
# <a name="override-element"></a><span data-ttu-id="d8467-102">Override 元素</span><span class="sxs-lookup"><span data-stu-id="d8467-102">Override element</span></span>

<span data-ttu-id="d8467-103">提供一种为其他区域设置指定某设置的值的方法。</span><span class="sxs-lookup"><span data-stu-id="d8467-103">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="d8467-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="d8467-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d8467-105">语法</span><span class="sxs-lookup"><span data-stu-id="d8467-105">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="d8467-106">包含于</span><span class="sxs-lookup"><span data-stu-id="d8467-106">Contained in</span></span>

|<span data-ttu-id="d8467-107">**Element**</span><span class="sxs-lookup"><span data-stu-id="d8467-107">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="d8467-108">CitationText</span><span class="sxs-lookup"><span data-stu-id="d8467-108">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="d8467-109">说明</span><span class="sxs-lookup"><span data-stu-id="d8467-109">Description</span></span>](description.md)|
|[<span data-ttu-id="d8467-110">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="d8467-110">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="d8467-111">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="d8467-111">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="d8467-112">DisplayName</span><span class="sxs-lookup"><span data-stu-id="d8467-112">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="d8467-113">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="d8467-113">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="d8467-114">IconUrl</span><span class="sxs-lookup"><span data-stu-id="d8467-114">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="d8467-115">QueryUri</span><span class="sxs-lookup"><span data-stu-id="d8467-115">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="d8467-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d8467-116">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="d8467-117">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="d8467-117">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="d8467-118">属性</span><span class="sxs-lookup"><span data-stu-id="d8467-118">Attributes</span></span>

|<span data-ttu-id="d8467-119">**属性**</span><span class="sxs-lookup"><span data-stu-id="d8467-119">**Attribute**</span></span>|<span data-ttu-id="d8467-120">**类型**</span><span class="sxs-lookup"><span data-stu-id="d8467-120">**Type**</span></span>|<span data-ttu-id="d8467-121">**必需**</span><span class="sxs-lookup"><span data-stu-id="d8467-121">**Required**</span></span>|<span data-ttu-id="d8467-122">**描述**</span><span class="sxs-lookup"><span data-stu-id="d8467-122">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d8467-123">区域设置</span><span class="sxs-lookup"><span data-stu-id="d8467-123">Locale</span></span>|<span data-ttu-id="d8467-124">string</span><span class="sxs-lookup"><span data-stu-id="d8467-124">string</span></span>|<span data-ttu-id="d8467-125">必需</span><span class="sxs-lookup"><span data-stu-id="d8467-125">required</span></span>|<span data-ttu-id="d8467-126">为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。</span><span class="sxs-lookup"><span data-stu-id="d8467-126">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="d8467-127">值</span><span class="sxs-lookup"><span data-stu-id="d8467-127">Value</span></span>|<span data-ttu-id="d8467-128">字符串</span><span class="sxs-lookup"><span data-stu-id="d8467-128">string</span></span>|<span data-ttu-id="d8467-129">必需</span><span class="sxs-lookup"><span data-stu-id="d8467-129">required</span></span>|<span data-ttu-id="d8467-130">指定表示为指定区域设置的设置的值。</span><span class="sxs-lookup"><span data-stu-id="d8467-130">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="d8467-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d8467-131">See also</span></span>

- [<span data-ttu-id="d8467-132">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="d8467-132">Localization for Office Add-ins</span></span>](/office/dev/add-ins/develop/localization)
    

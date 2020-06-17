---
title: 清单文件中的 SupportUrl 元素
description: SupportUrl 元素指定为外接程序提供支持信息的页面的 URL。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f75ee811699823a501ac594e66daaaf3f93c2782
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608703"
---
# <a name="supporturl-element"></a><span data-ttu-id="1057b-103">SupportUrl 元素</span><span class="sxs-lookup"><span data-stu-id="1057b-103">SupportUrl element</span></span>

<span data-ttu-id="1057b-104">指定提供外接程序支持信息的页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="1057b-104">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="1057b-105">语法</span><span class="sxs-lookup"><span data-stu-id="1057b-105">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="1057b-106">包含于</span><span class="sxs-lookup"><span data-stu-id="1057b-106">Contained in</span></span>

[<span data-ttu-id="1057b-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="1057b-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="1057b-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="1057b-108">Can contain</span></span>

|  <span data-ttu-id="1057b-109">元素</span><span class="sxs-lookup"><span data-stu-id="1057b-109">Element</span></span> | <span data-ttu-id="1057b-110">必需</span><span class="sxs-lookup"><span data-stu-id="1057b-110">Required</span></span> | <span data-ttu-id="1057b-111">Description</span><span class="sxs-lookup"><span data-stu-id="1057b-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1057b-112">Override</span><span class="sxs-lookup"><span data-stu-id="1057b-112">Override</span></span>](override.md)   | <span data-ttu-id="1057b-113">否</span><span class="sxs-lookup"><span data-stu-id="1057b-113">No</span></span> | <span data-ttu-id="1057b-114">指定其他区域设置 URL 的设置</span><span class="sxs-lookup"><span data-stu-id="1057b-114">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="1057b-115">属性</span><span class="sxs-lookup"><span data-stu-id="1057b-115">Attributes</span></span>

|<span data-ttu-id="1057b-116">**属性**</span><span class="sxs-lookup"><span data-stu-id="1057b-116">**Attribute**</span></span>|<span data-ttu-id="1057b-117">**类型**</span><span class="sxs-lookup"><span data-stu-id="1057b-117">**Type**</span></span>|<span data-ttu-id="1057b-118">**必需**</span><span class="sxs-lookup"><span data-stu-id="1057b-118">**Required**</span></span>|<span data-ttu-id="1057b-119">**描述**</span><span class="sxs-lookup"><span data-stu-id="1057b-119">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="1057b-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="1057b-120">DefaultValue</span></span>|<span data-ttu-id="1057b-121">URL</span><span class="sxs-lookup"><span data-stu-id="1057b-121">URL</span></span>|<span data-ttu-id="1057b-122">必需</span><span class="sxs-lookup"><span data-stu-id="1057b-122">required</span></span>|<span data-ttu-id="1057b-123">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="1057b-123">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

---
title: 清单文件中的 SupportUrl 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 18b9b7c4df9def70ab42ae213066188ac04c07a7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450413"
---
# <a name="supporturl-element"></a><span data-ttu-id="b43f2-102">SupportUrl 元素</span><span class="sxs-lookup"><span data-stu-id="b43f2-102">SupportUrl element</span></span>

<span data-ttu-id="b43f2-103">指定提供外接程序支持信息的页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="b43f2-103">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="b43f2-104">语法</span><span class="sxs-lookup"><span data-stu-id="b43f2-104">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="b43f2-105">包含于</span><span class="sxs-lookup"><span data-stu-id="b43f2-105">Contained in</span></span>

[<span data-ttu-id="b43f2-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b43f2-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="b43f2-107">可以包含</span><span class="sxs-lookup"><span data-stu-id="b43f2-107">Can contain</span></span>

|  <span data-ttu-id="b43f2-108">元素</span><span class="sxs-lookup"><span data-stu-id="b43f2-108">Element</span></span> | <span data-ttu-id="b43f2-109">必需</span><span class="sxs-lookup"><span data-stu-id="b43f2-109">Required</span></span> | <span data-ttu-id="b43f2-110">说明</span><span class="sxs-lookup"><span data-stu-id="b43f2-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b43f2-111">Override</span><span class="sxs-lookup"><span data-stu-id="b43f2-111">Override</span></span>](override.md)   | <span data-ttu-id="b43f2-112">否</span><span class="sxs-lookup"><span data-stu-id="b43f2-112">No</span></span> | <span data-ttu-id="b43f2-113">指定其他区域设置 URL 的设置</span><span class="sxs-lookup"><span data-stu-id="b43f2-113">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="b43f2-114">属性</span><span class="sxs-lookup"><span data-stu-id="b43f2-114">Attributes</span></span>

|<span data-ttu-id="b43f2-115">**属性**</span><span class="sxs-lookup"><span data-stu-id="b43f2-115">**Attribute**</span></span>|<span data-ttu-id="b43f2-116">**类型**</span><span class="sxs-lookup"><span data-stu-id="b43f2-116">**Type**</span></span>|<span data-ttu-id="b43f2-117">**必需**</span><span class="sxs-lookup"><span data-stu-id="b43f2-117">**Required**</span></span>|<span data-ttu-id="b43f2-118">**描述**</span><span class="sxs-lookup"><span data-stu-id="b43f2-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="b43f2-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="b43f2-119">DefaultValue</span></span>|<span data-ttu-id="b43f2-120">URL</span><span class="sxs-lookup"><span data-stu-id="b43f2-120">URL</span></span>|<span data-ttu-id="b43f2-121">必需</span><span class="sxs-lookup"><span data-stu-id="b43f2-121">required</span></span>|<span data-ttu-id="b43f2-122">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="b43f2-122">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

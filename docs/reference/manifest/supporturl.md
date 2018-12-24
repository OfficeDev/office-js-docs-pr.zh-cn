---
title: 清单文件中的 SupportUrl 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 00234ef9fe8960b9956e6a2595e2e2e71bfb97c6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432667"
---
# <a name="supporturl-element"></a><span data-ttu-id="db870-102">SupportUrl 元素</span><span class="sxs-lookup"><span data-stu-id="db870-102">SupportUrl element</span></span>

<span data-ttu-id="db870-103">指定提供外接程序支持信息的页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="db870-103">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="db870-104">语法</span><span class="sxs-lookup"><span data-stu-id="db870-104">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="db870-105">包含于</span><span class="sxs-lookup"><span data-stu-id="db870-105">Contained in</span></span>

[<span data-ttu-id="db870-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="db870-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="db870-107">可以包含</span><span class="sxs-lookup"><span data-stu-id="db870-107">Can contain</span></span>

|  <span data-ttu-id="db870-108">元素</span><span class="sxs-lookup"><span data-stu-id="db870-108">Element</span></span> | <span data-ttu-id="db870-109">必需</span><span class="sxs-lookup"><span data-stu-id="db870-109">Required</span></span> | <span data-ttu-id="db870-110">说明</span><span class="sxs-lookup"><span data-stu-id="db870-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="db870-111">Override</span><span class="sxs-lookup"><span data-stu-id="db870-111">Override</span></span>](override.md)   | <span data-ttu-id="db870-112">否</span><span class="sxs-lookup"><span data-stu-id="db870-112">No</span></span> | <span data-ttu-id="db870-113">指定其他区域设置 URL 的设置</span><span class="sxs-lookup"><span data-stu-id="db870-113">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="db870-114">属性</span><span class="sxs-lookup"><span data-stu-id="db870-114">Attributes</span></span>

|<span data-ttu-id="db870-115">**属性**</span><span class="sxs-lookup"><span data-stu-id="db870-115">**Attribute**</span></span>|<span data-ttu-id="db870-116">**类型**</span><span class="sxs-lookup"><span data-stu-id="db870-116">**Type**</span></span>|<span data-ttu-id="db870-117">**必需**</span><span class="sxs-lookup"><span data-stu-id="db870-117">**Required**</span></span>|<span data-ttu-id="db870-118">**说明**</span><span class="sxs-lookup"><span data-stu-id="db870-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="db870-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="db870-119">DefaultValue</span></span>|<span data-ttu-id="db870-120">URL</span><span class="sxs-lookup"><span data-stu-id="db870-120">URL</span></span>|<span data-ttu-id="db870-121">必需</span><span class="sxs-lookup"><span data-stu-id="db870-121">required</span></span>|<span data-ttu-id="db870-122">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="db870-122">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

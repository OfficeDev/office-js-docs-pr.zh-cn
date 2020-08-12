---
title: 清单文件中的 SupportUrl 元素
description: SupportUrl 元素指定为外接程序提供支持信息的页面的 URL。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: be516fe5848d775dacb0d424a92be02d59f85512
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641408"
---
# <a name="supporturl-element"></a><span data-ttu-id="f5b58-103">SupportUrl 元素</span><span class="sxs-lookup"><span data-stu-id="f5b58-103">SupportUrl element</span></span>

<span data-ttu-id="f5b58-104">指定提供外接程序支持信息的页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="f5b58-104">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="f5b58-105">语法</span><span class="sxs-lookup"><span data-stu-id="f5b58-105">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="f5b58-106">包含于</span><span class="sxs-lookup"><span data-stu-id="f5b58-106">Contained in</span></span>

[<span data-ttu-id="f5b58-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f5b58-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="f5b58-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="f5b58-108">Can contain</span></span>

|  <span data-ttu-id="f5b58-109">元素</span><span class="sxs-lookup"><span data-stu-id="f5b58-109">Element</span></span> | <span data-ttu-id="f5b58-110">必需</span><span class="sxs-lookup"><span data-stu-id="f5b58-110">Required</span></span> | <span data-ttu-id="f5b58-111">说明</span><span class="sxs-lookup"><span data-stu-id="f5b58-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f5b58-112">Override</span><span class="sxs-lookup"><span data-stu-id="f5b58-112">Override</span></span>](override.md)   | <span data-ttu-id="f5b58-113">否</span><span class="sxs-lookup"><span data-stu-id="f5b58-113">No</span></span> | <span data-ttu-id="f5b58-114">指定其他区域设置 URL 的设置</span><span class="sxs-lookup"><span data-stu-id="f5b58-114">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="f5b58-115">属性</span><span class="sxs-lookup"><span data-stu-id="f5b58-115">Attributes</span></span>

|<span data-ttu-id="f5b58-116">属性</span><span class="sxs-lookup"><span data-stu-id="f5b58-116">Attribute</span></span>|<span data-ttu-id="f5b58-117">类型</span><span class="sxs-lookup"><span data-stu-id="f5b58-117">Type</span></span>|<span data-ttu-id="f5b58-118">必需</span><span class="sxs-lookup"><span data-stu-id="f5b58-118">Required</span></span>|<span data-ttu-id="f5b58-119">说明</span><span class="sxs-lookup"><span data-stu-id="f5b58-119">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f5b58-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="f5b58-120">DefaultValue</span></span>|<span data-ttu-id="f5b58-121">URL</span><span class="sxs-lookup"><span data-stu-id="f5b58-121">URL</span></span>|<span data-ttu-id="f5b58-122">必需</span><span class="sxs-lookup"><span data-stu-id="f5b58-122">required</span></span>|<span data-ttu-id="f5b58-123">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="f5b58-123">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

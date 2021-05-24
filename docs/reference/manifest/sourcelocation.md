---
title: 清单文件中的 SourceLocation 元素
description: SourceLocation 元素指定外接程序的Office位置。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 4dcd093db2f23220eaa34c0c81300c4994c1a697
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590895"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="08cf3-103">SourceLocation 元素</span><span class="sxs-lookup"><span data-stu-id="08cf3-103">SourceLocation element</span></span>

<span data-ttu-id="08cf3-104">指定外接程序的源文件位置Office 1 到 2018 个字符之间的 URL。</span><span class="sxs-lookup"><span data-stu-id="08cf3-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="08cf3-105">源位置必须是 HTTPS 地址，而非文件路径。</span><span class="sxs-lookup"><span data-stu-id="08cf3-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="08cf3-106">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="08cf3-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="08cf3-107">语法</span><span class="sxs-lookup"><span data-stu-id="08cf3-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="08cf3-108">包含于</span><span class="sxs-lookup"><span data-stu-id="08cf3-108">Contained in</span></span>

- <span data-ttu-id="08cf3-109">[DefaultSettings](defaultsettings.md)（内容和任务窗格外接程序）</span><span class="sxs-lookup"><span data-stu-id="08cf3-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="08cf3-110">[FormSettings](formsettings.md)（邮件外接程序）</span><span class="sxs-lookup"><span data-stu-id="08cf3-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="08cf3-111">[ExtensionPoint](extensionpoint.md) (上下文和 LaunchEvent 邮件外接程序) </span><span class="sxs-lookup"><span data-stu-id="08cf3-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="08cf3-112">可以包含</span><span class="sxs-lookup"><span data-stu-id="08cf3-112">Can contain</span></span>

[<span data-ttu-id="08cf3-113">Override</span><span class="sxs-lookup"><span data-stu-id="08cf3-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="08cf3-114">属性</span><span class="sxs-lookup"><span data-stu-id="08cf3-114">Attributes</span></span>

|<span data-ttu-id="08cf3-115">属性</span><span class="sxs-lookup"><span data-stu-id="08cf3-115">Attribute</span></span>|<span data-ttu-id="08cf3-116">类型</span><span class="sxs-lookup"><span data-stu-id="08cf3-116">Type</span></span>|<span data-ttu-id="08cf3-117">必需</span><span class="sxs-lookup"><span data-stu-id="08cf3-117">Required</span></span>|<span data-ttu-id="08cf3-118">说明</span><span class="sxs-lookup"><span data-stu-id="08cf3-118">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="08cf3-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="08cf3-119">DefaultValue</span></span>|<span data-ttu-id="08cf3-120">URL</span><span class="sxs-lookup"><span data-stu-id="08cf3-120">URL</span></span>|<span data-ttu-id="08cf3-121">必需</span><span class="sxs-lookup"><span data-stu-id="08cf3-121">required</span></span>|<span data-ttu-id="08cf3-122">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="08cf3-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

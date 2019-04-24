---
title: 清单文件中的 SourceLocation 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7544e2bae480b9431c8912533ea1b761132a355e
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451974"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="9e8fe-102">SourceLocation 元素</span><span class="sxs-lookup"><span data-stu-id="9e8fe-102">SourceLocation element</span></span>

<span data-ttu-id="9e8fe-p101">指定 Office 外接程序的源文件位置为介于 1 和 2018 个字符之间的 URL。源位置必须是 HTTPS 地址，而非文件路径。</span><span class="sxs-lookup"><span data-stu-id="9e8fe-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="9e8fe-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="9e8fe-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9e8fe-106">语法</span><span class="sxs-lookup"><span data-stu-id="9e8fe-106">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="9e8fe-107">包含于</span><span class="sxs-lookup"><span data-stu-id="9e8fe-107">Contained in</span></span>

- <span data-ttu-id="9e8fe-108">[DefaultSettings](defaultsettings.md)（内容和任务窗格外接程序）</span><span class="sxs-lookup"><span data-stu-id="9e8fe-108">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="9e8fe-109">[FormSettings](formsettings.md)（邮件外接程序）</span><span class="sxs-lookup"><span data-stu-id="9e8fe-109">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="9e8fe-110">[ExtensionPoint](extensionpoint.md)（上下文邮件外接程序）</span><span class="sxs-lookup"><span data-stu-id="9e8fe-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="9e8fe-111">可以包含</span><span class="sxs-lookup"><span data-stu-id="9e8fe-111">Can contain</span></span>

[<span data-ttu-id="9e8fe-112">Override</span><span class="sxs-lookup"><span data-stu-id="9e8fe-112">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="9e8fe-113">属性</span><span class="sxs-lookup"><span data-stu-id="9e8fe-113">Attributes</span></span>

|<span data-ttu-id="9e8fe-114">**属性**</span><span class="sxs-lookup"><span data-stu-id="9e8fe-114">**Attribute**</span></span>|<span data-ttu-id="9e8fe-115">**类型**</span><span class="sxs-lookup"><span data-stu-id="9e8fe-115">**Type**</span></span>|<span data-ttu-id="9e8fe-116">**必需**</span><span class="sxs-lookup"><span data-stu-id="9e8fe-116">**Required**</span></span>|<span data-ttu-id="9e8fe-117">**描述**</span><span class="sxs-lookup"><span data-stu-id="9e8fe-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="9e8fe-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="9e8fe-118">DefaultValue</span></span>|<span data-ttu-id="9e8fe-119">URL</span><span class="sxs-lookup"><span data-stu-id="9e8fe-119">URL</span></span>|<span data-ttu-id="9e8fe-120">必需</span><span class="sxs-lookup"><span data-stu-id="9e8fe-120">required</span></span>|<span data-ttu-id="9e8fe-121">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="9e8fe-121">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

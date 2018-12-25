---
title: 清单文件中的 SourceLocation 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dc432ebb9482e8e9b8be5d90a838357ccf519ad3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433514"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="81d54-102">SourceLocation 元素</span><span class="sxs-lookup"><span data-stu-id="81d54-102">SourceLocation element</span></span>

<span data-ttu-id="81d54-p101">指定 Office 外接程序的源文件位置为介于 1 和 2018 个字符之间的 URL。源位置必须是 HTTPS 地址，而非文件路径。</span><span class="sxs-lookup"><span data-stu-id="81d54-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="81d54-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="81d54-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="81d54-106">语法</span><span class="sxs-lookup"><span data-stu-id="81d54-106">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="81d54-107">包含于</span><span class="sxs-lookup"><span data-stu-id="81d54-107">Contained in</span></span>

- <span data-ttu-id="81d54-108">[DefaultSettings](defaultsettings.md)（内容和任务窗格外接程序）</span><span class="sxs-lookup"><span data-stu-id="81d54-108">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="81d54-109">[FormSettings](formsettings.md)（邮件外接程序）</span><span class="sxs-lookup"><span data-stu-id="81d54-109">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="81d54-110">[ExtensionPoint](extensionpoint.md)（上下文邮件外接程序）</span><span class="sxs-lookup"><span data-stu-id="81d54-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="81d54-111">可以包含</span><span class="sxs-lookup"><span data-stu-id="81d54-111">Can contain</span></span>

[<span data-ttu-id="81d54-112">替代</span><span class="sxs-lookup"><span data-stu-id="81d54-112">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="81d54-113">属性</span><span class="sxs-lookup"><span data-stu-id="81d54-113">Attributes</span></span>

|<span data-ttu-id="81d54-114">**属性**</span><span class="sxs-lookup"><span data-stu-id="81d54-114">**Attribute**</span></span>|<span data-ttu-id="81d54-115">**类型**</span><span class="sxs-lookup"><span data-stu-id="81d54-115">**Type**</span></span>|<span data-ttu-id="81d54-116">**必需**</span><span class="sxs-lookup"><span data-stu-id="81d54-116">**Required**</span></span>|<span data-ttu-id="81d54-117">**说明**</span><span class="sxs-lookup"><span data-stu-id="81d54-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="81d54-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="81d54-118">DefaultValue</span></span>|<span data-ttu-id="81d54-119">URL</span><span class="sxs-lookup"><span data-stu-id="81d54-119">URL</span></span>|<span data-ttu-id="81d54-120">必需</span><span class="sxs-lookup"><span data-stu-id="81d54-120">required</span></span>|<span data-ttu-id="81d54-121">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="81d54-121">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

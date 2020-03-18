---
title: 清单文件中的 SourceLocation 元素
description: SourceLocation 元素指定 Office 外接程序的源文件位置。
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: fcca051b0d85c98cb011d5b886981c543ef8e3b0
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717899"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="1ccdb-103">SourceLocation 元素</span><span class="sxs-lookup"><span data-stu-id="1ccdb-103">SourceLocation element</span></span>

<span data-ttu-id="1ccdb-104">将 Office 外接程序的源文件位置指定为一个长度介于1到2018个字符之间的 URL。</span><span class="sxs-lookup"><span data-stu-id="1ccdb-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="1ccdb-105">源位置必须是 HTTPS 地址，而非文件路径。</span><span class="sxs-lookup"><span data-stu-id="1ccdb-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="1ccdb-106">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="1ccdb-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1ccdb-107">语法</span><span class="sxs-lookup"><span data-stu-id="1ccdb-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="1ccdb-108">包含于</span><span class="sxs-lookup"><span data-stu-id="1ccdb-108">Contained in</span></span>

- <span data-ttu-id="1ccdb-109">[DefaultSettings](defaultsettings.md)（内容和任务窗格外接程序）</span><span class="sxs-lookup"><span data-stu-id="1ccdb-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="1ccdb-110">[FormSettings](formsettings.md)（邮件外接程序）</span><span class="sxs-lookup"><span data-stu-id="1ccdb-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="1ccdb-111">[ExtensionPoint](extensionpoint.md)（上下文邮件外接程序）</span><span class="sxs-lookup"><span data-stu-id="1ccdb-111">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="1ccdb-112">可以包含</span><span class="sxs-lookup"><span data-stu-id="1ccdb-112">Can contain</span></span>

[<span data-ttu-id="1ccdb-113">Override</span><span class="sxs-lookup"><span data-stu-id="1ccdb-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="1ccdb-114">属性</span><span class="sxs-lookup"><span data-stu-id="1ccdb-114">Attributes</span></span>

|<span data-ttu-id="1ccdb-115">**属性**</span><span class="sxs-lookup"><span data-stu-id="1ccdb-115">**Attribute**</span></span>|<span data-ttu-id="1ccdb-116">**类型**</span><span class="sxs-lookup"><span data-stu-id="1ccdb-116">**Type**</span></span>|<span data-ttu-id="1ccdb-117">**必需**</span><span class="sxs-lookup"><span data-stu-id="1ccdb-117">**Required**</span></span>|<span data-ttu-id="1ccdb-118">**描述**</span><span class="sxs-lookup"><span data-stu-id="1ccdb-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="1ccdb-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="1ccdb-119">DefaultValue</span></span>|<span data-ttu-id="1ccdb-120">URL</span><span class="sxs-lookup"><span data-stu-id="1ccdb-120">URL</span></span>|<span data-ttu-id="1ccdb-121">必需</span><span class="sxs-lookup"><span data-stu-id="1ccdb-121">required</span></span>|<span data-ttu-id="1ccdb-122">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="1ccdb-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

---
title: 清单文件中的 AlternateId 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7a19715fa987978a4540b717f1d30fbff97157c5
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450665"
---
# <a name="alternateid-element"></a><span data-ttu-id="08bed-102">AlternateId 元素</span><span class="sxs-lookup"><span data-stu-id="08bed-102">AlternateId element</span></span>

<span data-ttu-id="08bed-103">指定由 Office 应用商店发布的 Office 外接程序的备用 ID。</span><span class="sxs-lookup"><span data-stu-id="08bed-103">Specifies the alternate ID for your Office Add-in as issued by the Office Store.</span></span>

<span data-ttu-id="08bed-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="08bed-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="08bed-105">语法</span><span class="sxs-lookup"><span data-stu-id="08bed-105">Syntax</span></span>

```XML
<AlternateId>string </AlternateId>
```

## <a name="contained-in"></a><span data-ttu-id="08bed-106">包含于</span><span class="sxs-lookup"><span data-stu-id="08bed-106">Contained in</span></span>

[<span data-ttu-id="08bed-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="08bed-107">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="08bed-108">注解</span><span class="sxs-lookup"><span data-stu-id="08bed-108">Remarks</span></span>

<span data-ttu-id="08bed-109">您不必自己创建此值；当您将外接程序提交至 Office 应用商店时，系统会为此外接程序分配该值。</span><span class="sxs-lookup"><span data-stu-id="08bed-109">You don't create this value yourself; it is assigned to your add-in when you submit it to the Office Store.</span></span>


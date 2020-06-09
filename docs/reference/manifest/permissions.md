---
title: 清单文件中的 Permissions 元素
description: 权限元素指定 Office 外接程序的 API 访问级别。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 603494b61ef126b35cb5cdff8c5f5b911bd25840
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611489"
---
# <a name="permissions-element"></a><span data-ttu-id="85337-103">Permissions 元素</span><span class="sxs-lookup"><span data-stu-id="85337-103">Permissions element</span></span>

<span data-ttu-id="85337-104">指定 Office 外接程序的 API 访问级别；您应基于最少特权的原则请求权限。</span><span class="sxs-lookup"><span data-stu-id="85337-104">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="85337-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="85337-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="85337-106">语法</span><span class="sxs-lookup"><span data-stu-id="85337-106">Syntax</span></span>

<span data-ttu-id="85337-107">对于内容和任务窗格外接程序：</span><span class="sxs-lookup"><span data-stu-id="85337-107">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="85337-108">对于邮件外接程序</span><span class="sxs-lookup"><span data-stu-id="85337-108">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="85337-109">包含于</span><span class="sxs-lookup"><span data-stu-id="85337-109">Contained in</span></span>

[<span data-ttu-id="85337-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="85337-110">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="85337-111">注解</span><span class="sxs-lookup"><span data-stu-id="85337-111">Remarks</span></span>

<span data-ttu-id="85337-112">有关详细信息，请参阅在[外接程序中请求 API 的使用权限](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)和[了解 Outlook 外接程序权限](../../outlook/understanding-outlook-add-in-permissions.md)。</span><span class="sxs-lookup"><span data-stu-id="85337-112">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>

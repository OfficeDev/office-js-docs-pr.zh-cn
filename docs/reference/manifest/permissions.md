---
title: 清单文件中的 Permissions 元素
description: 权限元素指定 Office 外接程序的 API 访问级别。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 91e024a2f13ea7605941c8c17a642f325cbcd61d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717997"
---
# <a name="permissions-element"></a><span data-ttu-id="c309c-103">Permissions 元素</span><span class="sxs-lookup"><span data-stu-id="c309c-103">Permissions element</span></span>

<span data-ttu-id="c309c-104">指定 Office 外接程序的 API 访问级别；您应基于最少特权的原则请求权限。</span><span class="sxs-lookup"><span data-stu-id="c309c-104">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="c309c-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="c309c-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c309c-106">语法</span><span class="sxs-lookup"><span data-stu-id="c309c-106">Syntax</span></span>

<span data-ttu-id="c309c-107">对于内容和任务窗格外接程序：</span><span class="sxs-lookup"><span data-stu-id="c309c-107">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="c309c-108">对于邮件外接程序</span><span class="sxs-lookup"><span data-stu-id="c309c-108">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="c309c-109">包含于</span><span class="sxs-lookup"><span data-stu-id="c309c-109">Contained in</span></span>

[<span data-ttu-id="c309c-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c309c-110">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="c309c-111">注解</span><span class="sxs-lookup"><span data-stu-id="c309c-111">Remarks</span></span>

<span data-ttu-id="c309c-112">有关详细信息，请参阅在[外接程序中请求 API 的使用权限](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)和[了解 Outlook 外接程序权限](../../outlook/understanding-outlook-add-in-permissions.md)。</span><span class="sxs-lookup"><span data-stu-id="c309c-112">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>

---
title: 清单文件中的 Permissions 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3442a8e0caee442ce1b38c5ff39cfd1ef5088fb7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450658"
---
# <a name="permissions-element"></a><span data-ttu-id="02401-102">Permissions 元素</span><span class="sxs-lookup"><span data-stu-id="02401-102">Permissions element</span></span>

<span data-ttu-id="02401-103">指定 Office 外接程序的 API 访问级别；您应基于最少特权的原则请求权限。</span><span class="sxs-lookup"><span data-stu-id="02401-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="02401-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="02401-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="02401-105">语法</span><span class="sxs-lookup"><span data-stu-id="02401-105">Syntax</span></span>

<span data-ttu-id="02401-106">对于内容和任务窗格外接程序：</span><span class="sxs-lookup"><span data-stu-id="02401-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="02401-107">对于邮件外接程序</span><span class="sxs-lookup"><span data-stu-id="02401-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="02401-108">包含于</span><span class="sxs-lookup"><span data-stu-id="02401-108">Contained in</span></span>

[<span data-ttu-id="02401-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="02401-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="02401-110">注解</span><span class="sxs-lookup"><span data-stu-id="02401-110">Remarks</span></span>

<span data-ttu-id="02401-111">有关详细信息，请参阅[在内容和任务窗格外接程序中请求 API 的使用权限](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)和[了解 Outlook 外接程序权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="02401-111">For more detail, see [Requesting permissions for API use in content and task pane add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

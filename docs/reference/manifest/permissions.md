---
title: 清单文件中的 Permissions 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a70d72e454273873c6a30ffd82c3a2a5194f55e0
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851304"
---
# <a name="permissions-element"></a><span data-ttu-id="a9db9-102">Permissions 元素</span><span class="sxs-lookup"><span data-stu-id="a9db9-102">Permissions element</span></span>

<span data-ttu-id="a9db9-103">指定 Office 外接程序的 API 访问级别；您应基于最少特权的原则请求权限。</span><span class="sxs-lookup"><span data-stu-id="a9db9-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="a9db9-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="a9db9-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a9db9-105">语法</span><span class="sxs-lookup"><span data-stu-id="a9db9-105">Syntax</span></span>

<span data-ttu-id="a9db9-106">对于内容和任务窗格外接程序：</span><span class="sxs-lookup"><span data-stu-id="a9db9-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="a9db9-107">对于邮件外接程序</span><span class="sxs-lookup"><span data-stu-id="a9db9-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="a9db9-108">包含于</span><span class="sxs-lookup"><span data-stu-id="a9db9-108">Contained in</span></span>

[<span data-ttu-id="a9db9-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="a9db9-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="a9db9-110">注解</span><span class="sxs-lookup"><span data-stu-id="a9db9-110">Remarks</span></span>

<span data-ttu-id="a9db9-111">有关详细信息，请参阅在[外接程序中请求 API 的使用权限](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)和[了解 Outlook 外接程序权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="a9db9-111">For more detail, see [Requesting permissions for API use in add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

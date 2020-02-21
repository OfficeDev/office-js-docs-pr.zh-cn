---
title: 清单文件中的 Permissions 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 95cb45f89e2a5b92edc29bf32d0b47fcb2dbf8ce
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165543"
---
# <a name="permissions-element"></a><span data-ttu-id="acf74-102">Permissions 元素</span><span class="sxs-lookup"><span data-stu-id="acf74-102">Permissions element</span></span>

<span data-ttu-id="acf74-103">指定 Office 外接程序的 API 访问级别；您应基于最少特权的原则请求权限。</span><span class="sxs-lookup"><span data-stu-id="acf74-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="acf74-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="acf74-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="acf74-105">语法</span><span class="sxs-lookup"><span data-stu-id="acf74-105">Syntax</span></span>

<span data-ttu-id="acf74-106">对于内容和任务窗格外接程序：</span><span class="sxs-lookup"><span data-stu-id="acf74-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="acf74-107">对于邮件外接程序</span><span class="sxs-lookup"><span data-stu-id="acf74-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="acf74-108">包含于</span><span class="sxs-lookup"><span data-stu-id="acf74-108">Contained in</span></span>

[<span data-ttu-id="acf74-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="acf74-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="acf74-110">注解</span><span class="sxs-lookup"><span data-stu-id="acf74-110">Remarks</span></span>

<span data-ttu-id="acf74-111">有关详细信息，请参阅在[外接程序中请求 API 的使用权限](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)和[了解 Outlook 外接程序权限](../../outlook/understanding-outlook-add-in-permissions.md)。</span><span class="sxs-lookup"><span data-stu-id="acf74-111">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>

---
title: 清单文件中的 Methods 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: b2ef9725b76b21af8d41b9e571d2851464aa1fcc
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596884"
---
# <a name="methods-element"></a><span data-ttu-id="a4984-102">Methods 元素</span><span class="sxs-lookup"><span data-stu-id="a4984-102">Methods element</span></span>

<span data-ttu-id="a4984-103">指定 Office 外接程序在激活时所需的 Office JavaScript API 方法的列表。</span><span class="sxs-lookup"><span data-stu-id="a4984-103">Specifies the list of Office JavaScript API methods that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="a4984-104">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="a4984-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="a4984-105">语法</span><span class="sxs-lookup"><span data-stu-id="a4984-105">Syntax</span></span>

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a><span data-ttu-id="a4984-106">包含于</span><span class="sxs-lookup"><span data-stu-id="a4984-106">Contained in</span></span>

[<span data-ttu-id="a4984-107">要求</span><span class="sxs-lookup"><span data-stu-id="a4984-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="a4984-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="a4984-108">Can contain</span></span>

[<span data-ttu-id="a4984-109">方法</span><span class="sxs-lookup"><span data-stu-id="a4984-109">Method</span></span>](method.md)

## <a name="remarks"></a><span data-ttu-id="a4984-110">注解</span><span class="sxs-lookup"><span data-stu-id="a4984-110">Remarks</span></span>

<span data-ttu-id="a4984-111">邮件外接程序中不支持**方法**和**方法**元素。有关要求集的详细信息，请参阅[Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="a4984-111">The **Methods** and **Method** elements aren't supported in mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

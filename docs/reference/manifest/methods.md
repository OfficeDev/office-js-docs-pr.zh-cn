---
title: 清单文件中的 Methods 元素
description: 方法元素指定 Office 外接程序在激活时所需的 Office JavaScript API 方法的列表。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d96eed07b6853cb51c24214b6017f14d6de6b83b
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718059"
---
# <a name="methods-element"></a><span data-ttu-id="a357f-103">Methods 元素</span><span class="sxs-lookup"><span data-stu-id="a357f-103">Methods element</span></span>

<span data-ttu-id="a357f-104">指定 Office 外接程序在激活时所需的 Office JavaScript API 方法的列表。</span><span class="sxs-lookup"><span data-stu-id="a357f-104">Specifies the list of Office JavaScript API methods that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="a357f-105">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="a357f-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="a357f-106">语法</span><span class="sxs-lookup"><span data-stu-id="a357f-106">Syntax</span></span>

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a><span data-ttu-id="a357f-107">包含于</span><span class="sxs-lookup"><span data-stu-id="a357f-107">Contained in</span></span>

[<span data-ttu-id="a357f-108">要求</span><span class="sxs-lookup"><span data-stu-id="a357f-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="a357f-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="a357f-109">Can contain</span></span>

[<span data-ttu-id="a357f-110">方法</span><span class="sxs-lookup"><span data-stu-id="a357f-110">Method</span></span>](method.md)

## <a name="remarks"></a><span data-ttu-id="a357f-111">注解</span><span class="sxs-lookup"><span data-stu-id="a357f-111">Remarks</span></span>

<span data-ttu-id="a357f-112">邮件外接程序中不支持**方法**和**方法**元素。有关要求集的详细信息，请参阅[Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="a357f-112">The **Methods** and **Method** elements aren't supported in mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

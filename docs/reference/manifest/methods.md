---
title: 清单文件中的 Methods 元素
description: 方法元素指定 Office 外接程序在激活时所需的 Office JavaScript API 方法的列表。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: b270122240314b792ee492336417a4d133bdcc84
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609018"
---
# <a name="methods-element"></a><span data-ttu-id="320b2-103">Methods 元素</span><span class="sxs-lookup"><span data-stu-id="320b2-103">Methods element</span></span>

<span data-ttu-id="320b2-104">指定 Office 外接程序在激活时所需的 Office JavaScript API 方法的列表。</span><span class="sxs-lookup"><span data-stu-id="320b2-104">Specifies the list of Office JavaScript API methods that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="320b2-105">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="320b2-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="320b2-106">语法</span><span class="sxs-lookup"><span data-stu-id="320b2-106">Syntax</span></span>

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a><span data-ttu-id="320b2-107">包含于</span><span class="sxs-lookup"><span data-stu-id="320b2-107">Contained in</span></span>

[<span data-ttu-id="320b2-108">要求</span><span class="sxs-lookup"><span data-stu-id="320b2-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="320b2-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="320b2-109">Can contain</span></span>

[<span data-ttu-id="320b2-110">方法</span><span class="sxs-lookup"><span data-stu-id="320b2-110">Method</span></span>](method.md)

## <a name="remarks"></a><span data-ttu-id="320b2-111">注解</span><span class="sxs-lookup"><span data-stu-id="320b2-111">Remarks</span></span>

<span data-ttu-id="320b2-112">邮件外接程序中不支持**方法**和**方法**元素。有关要求集的详细信息，请参阅[Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="320b2-112">The **Methods** and **Method** elements aren't supported in mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

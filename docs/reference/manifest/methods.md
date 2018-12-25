---
title: 清单文件中的 Methods 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 6e280cb49eadef587cd3a91e0664ece3c3d59f50
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432751"
---
# <a name="methods-element"></a><span data-ttu-id="21830-102">Methods 元素</span><span class="sxs-lookup"><span data-stu-id="21830-102">Methods element</span></span>

<span data-ttu-id="21830-103">指定适用于 Office 的 JavaScript API 的方法列表，Office 外接程序需要该方法列表才能激活。</span><span class="sxs-lookup"><span data-stu-id="21830-103">Specifies the list of JavaScript API for Office methods that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="21830-104">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="21830-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="21830-105">语法</span><span class="sxs-lookup"><span data-stu-id="21830-105">Syntax</span></span>

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a><span data-ttu-id="21830-106">包含于</span><span class="sxs-lookup"><span data-stu-id="21830-106">Contained in</span></span>

[<span data-ttu-id="21830-107">要求</span><span class="sxs-lookup"><span data-stu-id="21830-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="21830-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="21830-108">Can contain</span></span>

[<span data-ttu-id="21830-109">方法</span><span class="sxs-lookup"><span data-stu-id="21830-109">Method</span></span>](method.md)

## <a name="remarks"></a><span data-ttu-id="21830-110">注释</span><span class="sxs-lookup"><span data-stu-id="21830-110">Remarks</span></span>

<span data-ttu-id="21830-111">**Methods** 和 **Method** 元素不受邮件外接程序的支持。有关要求集的详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="21830-111">The  Methods and Method elements aren't supported in mail add-ins. For more information about requirement sets, see Specify Office hosts and API requirements.</span></span>


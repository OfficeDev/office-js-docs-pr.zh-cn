---
title: 清单文件中的 Version 元素
description: Version 元素指定 Office 外接程序版本。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 48a2be94d95ece597e47468bb18db2a7962a51e9
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173932"
---
# <a name="version-element"></a><span data-ttu-id="ca3b0-103">Version 元素</span><span class="sxs-lookup"><span data-stu-id="ca3b0-103">Version element</span></span>

<span data-ttu-id="ca3b0-104">指定 Office 外接程序的版本。</span><span class="sxs-lookup"><span data-stu-id="ca3b0-104">Specifies the version of your Office Add-in.</span></span> <span data-ttu-id="ca3b0-105">版本号可以是 1、2、3 或 4 部分 (例如 n、n.n、n.n 或 n.n.n) 。</span><span class="sxs-lookup"><span data-stu-id="ca3b0-105">The version number can be 1, 2, 3, or 4 parts (i.e., n, n.n, n.n.n, or n.n.n.n).</span></span>

<span data-ttu-id="ca3b0-106">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="ca3b0-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ca3b0-107">语法</span><span class="sxs-lookup"><span data-stu-id="ca3b0-107">Syntax</span></span>

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a><span data-ttu-id="ca3b0-108">包含于</span><span class="sxs-lookup"><span data-stu-id="ca3b0-108">Contained in</span></span>

[<span data-ttu-id="ca3b0-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="ca3b0-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="ca3b0-110">注解</span><span class="sxs-lookup"><span data-stu-id="ca3b0-110">Remarks</span></span>

<span data-ttu-id="ca3b0-111">版本号的每个部分最多为 5 位数字。</span><span class="sxs-lookup"><span data-stu-id="ca3b0-111">Each part of the version number can be a maximum of 5 digits.</span></span>

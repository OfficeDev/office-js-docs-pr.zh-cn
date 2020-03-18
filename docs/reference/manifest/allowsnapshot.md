---
title: 清单文件中的 AllowSnapshot 元素
description: 指定是否将内容外接程序的快照图像与主机文档一起保存。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8bb143d13a17b3e184af64f1bf18f2a32a55b60c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720958"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="189c6-103">AllowSnapshot 元素</span><span class="sxs-lookup"><span data-stu-id="189c6-103">AllowSnapshot element</span></span>

<span data-ttu-id="189c6-104">指定是否将内容外接程序的快照图像与主机文档一起保存。</span><span class="sxs-lookup"><span data-stu-id="189c6-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="189c6-105">**外接程序类型：** 内容</span><span class="sxs-lookup"><span data-stu-id="189c6-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="189c6-106">语法</span><span class="sxs-lookup"><span data-stu-id="189c6-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="189c6-107">包含于</span><span class="sxs-lookup"><span data-stu-id="189c6-107">Contained in</span></span>

[<span data-ttu-id="189c6-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="189c6-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="189c6-109">注释</span><span class="sxs-lookup"><span data-stu-id="189c6-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="189c6-110">**AllowSnapshot** 在默认情况下为 `true`。</span><span class="sxs-lookup"><span data-stu-id="189c6-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="189c6-111">这样，用户在不支持 Office 外接程序的主机应用程序版本中打开文档时，即可看到该外接程序的图像，或者如果主机应用程序无法连接到托管外接程序的服务器时，会提供该外接程序的静态图像。</span><span class="sxs-lookup"><span data-stu-id="189c6-111">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="189c6-112">但是，这也意味着可以直接从托管该外接程序的文档访问显示在外接程序中的潜在敏感信息。</span><span class="sxs-lookup"><span data-stu-id="189c6-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>


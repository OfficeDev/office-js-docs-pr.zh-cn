---
title: 清单文件中的 AllowSnapshot 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: f1aced0ce37b01c277ea5a8621f6c7764d2f761b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432345"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="b3310-102">AllowSnapshot 元素</span><span class="sxs-lookup"><span data-stu-id="b3310-102">AllowSnapshot element</span></span>

<span data-ttu-id="b3310-103">指定是否将内容外接程序的快照图像与主机文档一起保存。</span><span class="sxs-lookup"><span data-stu-id="b3310-103">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="b3310-104">**外接程序类型：** 内容</span><span class="sxs-lookup"><span data-stu-id="b3310-104">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="b3310-105">语法</span><span class="sxs-lookup"><span data-stu-id="b3310-105">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="b3310-106">包含于</span><span class="sxs-lookup"><span data-stu-id="b3310-106">Contained in</span></span>

[<span data-ttu-id="b3310-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b3310-107">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="b3310-108">注释</span><span class="sxs-lookup"><span data-stu-id="b3310-108">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="b3310-109">**AllowSnapshot** 在默认情况下为 `true`。</span><span class="sxs-lookup"><span data-stu-id="b3310-109">**Security Note:AllowSnapshot** is true`true` by default.</span></span> <span data-ttu-id="b3310-110">这样，用户在不支持 Office 外接程序的主机应用程序版本中打开文档时，即可看到该外接程序的图像，或者如果主机应用程序无法连接到托管外接程序的服务器时，会提供该外接程序的静态图像。</span><span class="sxs-lookup"><span data-stu-id="b3310-110">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="b3310-111">但是，这也意味着可以直接从托管该外接程序的文档访问显示在外接程序中的潜在敏感信息。</span><span class="sxs-lookup"><span data-stu-id="b3310-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>


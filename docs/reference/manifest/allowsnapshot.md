---
title: 清单文件中的 AllowSnapshot 元素
description: 指定是否将内容外接程序的快照图像与主机文档一起保存。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ea910e1ad747e304dbc6ab4fbdcf44a9610dab19
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294274"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="b0d1a-103">AllowSnapshot 元素</span><span class="sxs-lookup"><span data-stu-id="b0d1a-103">AllowSnapshot element</span></span>

<span data-ttu-id="b0d1a-104">指定是否将内容外接程序的快照图像与主机文档一起保存。</span><span class="sxs-lookup"><span data-stu-id="b0d1a-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="b0d1a-105">**外接程序类型：** 内容</span><span class="sxs-lookup"><span data-stu-id="b0d1a-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="b0d1a-106">语法</span><span class="sxs-lookup"><span data-stu-id="b0d1a-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="b0d1a-107">包含于</span><span class="sxs-lookup"><span data-stu-id="b0d1a-107">Contained in</span></span>

[<span data-ttu-id="b0d1a-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b0d1a-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="b0d1a-109">注释</span><span class="sxs-lookup"><span data-stu-id="b0d1a-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="b0d1a-110">**AllowSnapshot** 在默认情况下为 `true`。</span><span class="sxs-lookup"><span data-stu-id="b0d1a-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="b0d1a-111">这使得在不支持 Office 外接程序的 Office 应用程序版本中打开文档的用户可以看到加载项的图像，如果应用程序无法连接到承载外接程序的服务器，则会提供该外接程序的静态图像。</span><span class="sxs-lookup"><span data-stu-id="b0d1a-111">This makes an image of the add-in visible for users that open the document in a version of the Office application that doesn't support Office Add-ins, or provides a static image of the add-in if the application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="b0d1a-112">但是，这也意味着可以直接从托管该外接程序的文档访问显示在外接程序中的潜在敏感信息。</span><span class="sxs-lookup"><span data-stu-id="b0d1a-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

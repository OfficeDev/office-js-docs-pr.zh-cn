---
title: 清单文件中的 Hosts 元素
description: 指定将在其中激活 Office 外接程序的 Office 客户端应用程序。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cd4e0eecce610b10fdc9dafcde7b807fde425b14
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718102"
---
# <a name="hosts-element"></a><span data-ttu-id="19aa9-103">Hosts 元素</span><span class="sxs-lookup"><span data-stu-id="19aa9-103">Hosts element</span></span>

<span data-ttu-id="19aa9-p101">指定将在其中激活 Office 外接程序的 Office 客户端应用程序。包含 **Host** 元素及其设置的集合。</span><span class="sxs-lookup"><span data-stu-id="19aa9-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="19aa9-106">当该元素被包括在 [VersionOverrides](versionoverrides.md)(#versionoverrides) 节点中时，它将替代清单中父级部分中的 **Hosts** 元素。</span><span class="sxs-lookup"><span data-stu-id="19aa9-106">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="19aa9-107">子元素</span><span class="sxs-lookup"><span data-stu-id="19aa9-107">Child elements</span></span>

|  <span data-ttu-id="19aa9-108">元素</span><span class="sxs-lookup"><span data-stu-id="19aa9-108">Element</span></span> |  <span data-ttu-id="19aa9-109">必需</span><span class="sxs-lookup"><span data-stu-id="19aa9-109">Required</span></span>  |  <span data-ttu-id="19aa9-110">说明</span><span class="sxs-lookup"><span data-stu-id="19aa9-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="19aa9-111">Host</span><span class="sxs-lookup"><span data-stu-id="19aa9-111">Host</span></span>](host.md)    |  <span data-ttu-id="19aa9-112">是</span><span class="sxs-lookup"><span data-stu-id="19aa9-112">Yes</span></span>   |  <span data-ttu-id="19aa9-113">说明主机及其设置。</span><span class="sxs-lookup"><span data-stu-id="19aa9-113">Describes a host and its settings.</span></span> |

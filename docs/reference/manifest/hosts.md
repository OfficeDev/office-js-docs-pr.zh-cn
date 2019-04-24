---
title: 清单文件中的 Hosts 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 606073977366e37ecc4419f468f01bfb25647a7d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452023"
---
# <a name="hosts-element"></a><span data-ttu-id="8a10c-102">Hosts 元素</span><span class="sxs-lookup"><span data-stu-id="8a10c-102">Hosts element</span></span>

<span data-ttu-id="8a10c-p101">指定将在其中激活 Office 外接程序的 Office 客户端应用程序。包含 **Host** 元素及其设置的集合。</span><span class="sxs-lookup"><span data-stu-id="8a10c-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="8a10c-105">当该元素被包括在 [VersionOverrides](versionoverrides.md)(#versionoverrides) 节点中时，它将替代清单中父级部分中的 **Hosts** 元素。</span><span class="sxs-lookup"><span data-stu-id="8a10c-105">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="8a10c-106">子元素</span><span class="sxs-lookup"><span data-stu-id="8a10c-106">Child elements</span></span>

|  <span data-ttu-id="8a10c-107">元素</span><span class="sxs-lookup"><span data-stu-id="8a10c-107">Element</span></span> |  <span data-ttu-id="8a10c-108">必需</span><span class="sxs-lookup"><span data-stu-id="8a10c-108">Required</span></span>  |  <span data-ttu-id="8a10c-109">说明</span><span class="sxs-lookup"><span data-stu-id="8a10c-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8a10c-110">Host</span><span class="sxs-lookup"><span data-stu-id="8a10c-110">Host</span></span>](host.md)    |  <span data-ttu-id="8a10c-111">是</span><span class="sxs-lookup"><span data-stu-id="8a10c-111">Yes</span></span>   |  <span data-ttu-id="8a10c-112">说明主机及其设置。</span><span class="sxs-lookup"><span data-stu-id="8a10c-112">Describes a host and its settings.</span></span> |

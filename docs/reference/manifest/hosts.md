---
title: 清单文件中的 Hosts 元素
description: 指定将在其中激活 Office 外接程序的 Office 客户端应用程序。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 037ac2b5fedbfb1b59b7523382574942fe59a00a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611804"
---
# <a name="hosts-element"></a><span data-ttu-id="d2a51-103">Hosts 元素</span><span class="sxs-lookup"><span data-stu-id="d2a51-103">Hosts element</span></span>

<span data-ttu-id="d2a51-p101">指定将在其中激活 Office 外接程序的 Office 客户端应用程序。包含 **Host** 元素及其设置的集合。</span><span class="sxs-lookup"><span data-stu-id="d2a51-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="d2a51-106">当该元素被包括在 [VersionOverrides](versionoverrides.md)(#versionoverrides) 节点中时，它将替代清单中父级部分中的 **Hosts** 元素。</span><span class="sxs-lookup"><span data-stu-id="d2a51-106">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="d2a51-107">子元素</span><span class="sxs-lookup"><span data-stu-id="d2a51-107">Child elements</span></span>

|  <span data-ttu-id="d2a51-108">元素</span><span class="sxs-lookup"><span data-stu-id="d2a51-108">Element</span></span> |  <span data-ttu-id="d2a51-109">必需</span><span class="sxs-lookup"><span data-stu-id="d2a51-109">Required</span></span>  |  <span data-ttu-id="d2a51-110">Description</span><span class="sxs-lookup"><span data-stu-id="d2a51-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d2a51-111">Host</span><span class="sxs-lookup"><span data-stu-id="d2a51-111">Host</span></span>](host.md)    |  <span data-ttu-id="d2a51-112">是</span><span class="sxs-lookup"><span data-stu-id="d2a51-112">Yes</span></span>   |  <span data-ttu-id="d2a51-113">说明主机及其设置。</span><span class="sxs-lookup"><span data-stu-id="d2a51-113">Describes a host and its settings.</span></span> |

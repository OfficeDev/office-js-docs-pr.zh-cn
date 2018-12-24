---
title: 清单文件中的 Hosts 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 59010c0f6c0d14d8721856f81def11540db28704
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433409"
---
# <a name="hosts-element"></a><span data-ttu-id="3e6a7-102">Hosts 元素</span><span class="sxs-lookup"><span data-stu-id="3e6a7-102">Hosts element</span></span>

<span data-ttu-id="3e6a7-p101">指定将在其中激活 Office 外接程序的 Office 客户端应用程序。包含 **Host** 元素及其设置的集合。</span><span class="sxs-lookup"><span data-stu-id="3e6a7-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="3e6a7-105">当该元素被包括在 [VersionOverrides](versionoverrides.md)(#versionoverrides) 节点中时，它将替代清单中父级部分中的 **Hosts** 元素。</span><span class="sxs-lookup"><span data-stu-id="3e6a7-105">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="3e6a7-106">子元素</span><span class="sxs-lookup"><span data-stu-id="3e6a7-106">Child elements</span></span>

|  <span data-ttu-id="3e6a7-107">元素</span><span class="sxs-lookup"><span data-stu-id="3e6a7-107">Element</span></span> |  <span data-ttu-id="3e6a7-108">必需</span><span class="sxs-lookup"><span data-stu-id="3e6a7-108">Required</span></span>  |  <span data-ttu-id="3e6a7-109">说明</span><span class="sxs-lookup"><span data-stu-id="3e6a7-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3e6a7-110">Host</span><span class="sxs-lookup"><span data-stu-id="3e6a7-110">Host</span></span>](host.md)    |  <span data-ttu-id="3e6a7-111">是</span><span class="sxs-lookup"><span data-stu-id="3e6a7-111">Yes</span></span>   |  <span data-ttu-id="3e6a7-112">说明主机及其设置。</span><span class="sxs-lookup"><span data-stu-id="3e6a7-112">Describes a host and its settings.</span></span> |

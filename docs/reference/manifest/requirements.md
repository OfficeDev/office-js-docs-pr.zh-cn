---
title: 清单文件中的 Requirements 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 2544e9b01b2d4d3ddc0a0c6238b4a5b0e6c4f832
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432702"
---
# <a name="requirements-element"></a><span data-ttu-id="8499e-102">Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="8499e-102">Requirements element</span></span>

<span data-ttu-id="8499e-103">指定适用于 Office 的 JavaScript API 要求（[要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)和/或方法）的最小集，Office 外接程序需要该集才能激活。</span><span class="sxs-lookup"><span data-stu-id="8499e-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="8499e-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="8499e-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8499e-105">语法</span><span class="sxs-lookup"><span data-stu-id="8499e-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="8499e-106">包含于</span><span class="sxs-lookup"><span data-stu-id="8499e-106">Contained in</span></span>

[<span data-ttu-id="8499e-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="8499e-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="8499e-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="8499e-108">Can contain</span></span>

|<span data-ttu-id="8499e-109">**元素**</span><span class="sxs-lookup"><span data-stu-id="8499e-109">**Element**</span></span>|<span data-ttu-id="8499e-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="8499e-110">**Content**</span></span>|<span data-ttu-id="8499e-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="8499e-111">**Mail**</span></span>|<span data-ttu-id="8499e-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="8499e-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="8499e-113">Sets</span><span class="sxs-lookup"><span data-stu-id="8499e-113">Sets</span></span>](sets.md)|<span data-ttu-id="8499e-114">x</span><span class="sxs-lookup"><span data-stu-id="8499e-114">x</span></span>|<span data-ttu-id="8499e-115">x</span><span class="sxs-lookup"><span data-stu-id="8499e-115">x</span></span>|<span data-ttu-id="8499e-116">x</span><span class="sxs-lookup"><span data-stu-id="8499e-116">x</span></span>|
|[<span data-ttu-id="8499e-117">Methods</span><span class="sxs-lookup"><span data-stu-id="8499e-117">Methods</span></span>](methods.md)|<span data-ttu-id="8499e-118">x</span><span class="sxs-lookup"><span data-stu-id="8499e-118">x</span></span>||<span data-ttu-id="8499e-119">x</span><span class="sxs-lookup"><span data-stu-id="8499e-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="8499e-120">注释</span><span class="sxs-lookup"><span data-stu-id="8499e-120">Remarks</span></span>

<span data-ttu-id="8499e-121">有关要求集的详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="8499e-121">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>


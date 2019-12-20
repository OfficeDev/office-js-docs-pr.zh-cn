---
title: "\"Context.subname\"： \"邮箱\"。诊断-要求集1。7"
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 3baf192dc209d015ff888ff5067d2cafbaee3181
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814624"
---
# <a name="diagnostics"></a><span data-ttu-id="916c3-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="916c3-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="916c3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="916c3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="916c3-104">将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="916c3-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="916c3-105">要求</span><span class="sxs-lookup"><span data-stu-id="916c3-105">Requirements</span></span>

|<span data-ttu-id="916c3-106">要求</span><span class="sxs-lookup"><span data-stu-id="916c3-106">Requirement</span></span>| <span data-ttu-id="916c3-107">值</span><span class="sxs-lookup"><span data-stu-id="916c3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="916c3-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="916c3-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="916c3-109">1.1</span><span class="sxs-lookup"><span data-stu-id="916c3-109">1.1</span></span>|
|[<span data-ttu-id="916c3-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="916c3-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="916c3-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="916c3-111">ReadItem</span></span>|
|[<span data-ttu-id="916c3-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="916c3-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="916c3-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="916c3-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="916c3-114">属性</span><span class="sxs-lookup"><span data-stu-id="916c3-114">Properties</span></span>

| <span data-ttu-id="916c3-115">属性</span><span class="sxs-lookup"><span data-stu-id="916c3-115">Property</span></span> | <span data-ttu-id="916c3-116">最低</span><span class="sxs-lookup"><span data-stu-id="916c3-116">Minimum</span></span><br><span data-ttu-id="916c3-117">权限级别</span><span class="sxs-lookup"><span data-stu-id="916c3-117">permission level</span></span> | <span data-ttu-id="916c3-118">型号</span><span class="sxs-lookup"><span data-stu-id="916c3-118">Modes</span></span> | <span data-ttu-id="916c3-119">返回类型</span><span class="sxs-lookup"><span data-stu-id="916c3-119">Return type</span></span> | <span data-ttu-id="916c3-120">最低</span><span class="sxs-lookup"><span data-stu-id="916c3-120">Minimum</span></span><br><span data-ttu-id="916c3-121">要求集</span><span class="sxs-lookup"><span data-stu-id="916c3-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="916c3-122">主机名</span><span class="sxs-lookup"><span data-stu-id="916c3-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.7#hostname) | <span data-ttu-id="916c3-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="916c3-123">ReadItem</span></span> | <span data-ttu-id="916c3-124">撰写</span><span class="sxs-lookup"><span data-stu-id="916c3-124">Compose</span></span><br><span data-ttu-id="916c3-125">读取</span><span class="sxs-lookup"><span data-stu-id="916c3-125">Read</span></span> | <span data-ttu-id="916c3-126">String</span><span class="sxs-lookup"><span data-stu-id="916c3-126">String</span></span> | [<span data-ttu-id="916c3-127">1.1</span><span class="sxs-lookup"><span data-stu-id="916c3-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="916c3-128">Diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="916c3-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.7#hostversion) | <span data-ttu-id="916c3-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="916c3-129">ReadItem</span></span> | <span data-ttu-id="916c3-130">撰写</span><span class="sxs-lookup"><span data-stu-id="916c3-130">Compose</span></span><br><span data-ttu-id="916c3-131">读取</span><span class="sxs-lookup"><span data-stu-id="916c3-131">Read</span></span> | <span data-ttu-id="916c3-132">String</span><span class="sxs-lookup"><span data-stu-id="916c3-132">String</span></span> | [<span data-ttu-id="916c3-133">1.1</span><span class="sxs-lookup"><span data-stu-id="916c3-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="916c3-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="916c3-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.7#owaview) | <span data-ttu-id="916c3-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="916c3-135">ReadItem</span></span> | <span data-ttu-id="916c3-136">撰写</span><span class="sxs-lookup"><span data-stu-id="916c3-136">Compose</span></span><br><span data-ttu-id="916c3-137">读取</span><span class="sxs-lookup"><span data-stu-id="916c3-137">Read</span></span> | <span data-ttu-id="916c3-138">String</span><span class="sxs-lookup"><span data-stu-id="916c3-138">String</span></span> | [<span data-ttu-id="916c3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="916c3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

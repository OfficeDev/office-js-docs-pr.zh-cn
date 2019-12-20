---
title: Office.context.mailbox.diagnostics - 要求集 1.2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: dad9d35c397351938944d89bf98e450427cb74a3
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814981"
---
# <a name="diagnostics"></a><span data-ttu-id="09623-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="09623-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="09623-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="09623-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="09623-104">将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="09623-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09623-105">要求</span><span class="sxs-lookup"><span data-stu-id="09623-105">Requirements</span></span>

|<span data-ttu-id="09623-106">要求</span><span class="sxs-lookup"><span data-stu-id="09623-106">Requirement</span></span>| <span data-ttu-id="09623-107">值</span><span class="sxs-lookup"><span data-stu-id="09623-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="09623-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09623-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09623-109">1.1</span><span class="sxs-lookup"><span data-stu-id="09623-109">1.1</span></span>|
|[<span data-ttu-id="09623-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09623-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09623-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09623-111">ReadItem</span></span>|
|[<span data-ttu-id="09623-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09623-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09623-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="09623-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="09623-114">属性</span><span class="sxs-lookup"><span data-stu-id="09623-114">Properties</span></span>

| <span data-ttu-id="09623-115">属性</span><span class="sxs-lookup"><span data-stu-id="09623-115">Property</span></span> | <span data-ttu-id="09623-116">最低</span><span class="sxs-lookup"><span data-stu-id="09623-116">Minimum</span></span><br><span data-ttu-id="09623-117">权限级别</span><span class="sxs-lookup"><span data-stu-id="09623-117">permission level</span></span> | <span data-ttu-id="09623-118">型号</span><span class="sxs-lookup"><span data-stu-id="09623-118">Modes</span></span> | <span data-ttu-id="09623-119">返回类型</span><span class="sxs-lookup"><span data-stu-id="09623-119">Return type</span></span> | <span data-ttu-id="09623-120">最低</span><span class="sxs-lookup"><span data-stu-id="09623-120">Minimum</span></span><br><span data-ttu-id="09623-121">要求集</span><span class="sxs-lookup"><span data-stu-id="09623-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="09623-122">主机名</span><span class="sxs-lookup"><span data-stu-id="09623-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2#hostname) | <span data-ttu-id="09623-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09623-123">ReadItem</span></span> | <span data-ttu-id="09623-124">撰写</span><span class="sxs-lookup"><span data-stu-id="09623-124">Compose</span></span><br><span data-ttu-id="09623-125">读取</span><span class="sxs-lookup"><span data-stu-id="09623-125">Read</span></span> | <span data-ttu-id="09623-126">String</span><span class="sxs-lookup"><span data-stu-id="09623-126">String</span></span> | [<span data-ttu-id="09623-127">1.1</span><span class="sxs-lookup"><span data-stu-id="09623-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="09623-128">Diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="09623-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2#hostversion) | <span data-ttu-id="09623-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09623-129">ReadItem</span></span> | <span data-ttu-id="09623-130">撰写</span><span class="sxs-lookup"><span data-stu-id="09623-130">Compose</span></span><br><span data-ttu-id="09623-131">读取</span><span class="sxs-lookup"><span data-stu-id="09623-131">Read</span></span> | <span data-ttu-id="09623-132">String</span><span class="sxs-lookup"><span data-stu-id="09623-132">String</span></span> | [<span data-ttu-id="09623-133">1.1</span><span class="sxs-lookup"><span data-stu-id="09623-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="09623-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="09623-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2#owaview) | <span data-ttu-id="09623-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09623-135">ReadItem</span></span> | <span data-ttu-id="09623-136">撰写</span><span class="sxs-lookup"><span data-stu-id="09623-136">Compose</span></span><br><span data-ttu-id="09623-137">读取</span><span class="sxs-lookup"><span data-stu-id="09623-137">Read</span></span> | <span data-ttu-id="09623-138">String</span><span class="sxs-lookup"><span data-stu-id="09623-138">String</span></span> | [<span data-ttu-id="09623-139">1.1</span><span class="sxs-lookup"><span data-stu-id="09623-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

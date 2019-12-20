---
title: "\"Context.subname\"： \"邮箱. userProfile-要求集 1.5\""
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 6b5229c1bc300d11714f3aa2cf8fa8ff2465667c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814778"
---
# <a name="userprofile"></a><span data-ttu-id="af664-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="af664-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="af664-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="af664-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="af664-104">提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="af664-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="af664-105">要求</span><span class="sxs-lookup"><span data-stu-id="af664-105">Requirements</span></span>

|<span data-ttu-id="af664-106">要求</span><span class="sxs-lookup"><span data-stu-id="af664-106">Requirement</span></span>| <span data-ttu-id="af664-107">值</span><span class="sxs-lookup"><span data-stu-id="af664-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="af664-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="af664-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af664-109">1.1</span><span class="sxs-lookup"><span data-stu-id="af664-109">1.1</span></span>|
|[<span data-ttu-id="af664-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="af664-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af664-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af664-111">ReadItem</span></span>|
|[<span data-ttu-id="af664-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="af664-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af664-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="af664-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="af664-114">属性</span><span class="sxs-lookup"><span data-stu-id="af664-114">Properties</span></span>

| <span data-ttu-id="af664-115">属性</span><span class="sxs-lookup"><span data-stu-id="af664-115">Property</span></span> | <span data-ttu-id="af664-116">最低</span><span class="sxs-lookup"><span data-stu-id="af664-116">Minimum</span></span><br><span data-ttu-id="af664-117">权限级别</span><span class="sxs-lookup"><span data-stu-id="af664-117">permission level</span></span> | <span data-ttu-id="af664-118">型号</span><span class="sxs-lookup"><span data-stu-id="af664-118">Modes</span></span> | <span data-ttu-id="af664-119">返回类型</span><span class="sxs-lookup"><span data-stu-id="af664-119">Return type</span></span> | <span data-ttu-id="af664-120">最低</span><span class="sxs-lookup"><span data-stu-id="af664-120">Minimum</span></span><br><span data-ttu-id="af664-121">要求集</span><span class="sxs-lookup"><span data-stu-id="af664-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="af664-122">displayName</span><span class="sxs-lookup"><span data-stu-id="af664-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | <span data-ttu-id="af664-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af664-123">ReadItem</span></span> | <span data-ttu-id="af664-124">撰写</span><span class="sxs-lookup"><span data-stu-id="af664-124">Compose</span></span><br><span data-ttu-id="af664-125">读取</span><span class="sxs-lookup"><span data-stu-id="af664-125">Read</span></span> | <span data-ttu-id="af664-126">String</span><span class="sxs-lookup"><span data-stu-id="af664-126">String</span></span> | [<span data-ttu-id="af664-127">1.1</span><span class="sxs-lookup"><span data-stu-id="af664-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af664-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="af664-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | <span data-ttu-id="af664-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af664-129">ReadItem</span></span> | <span data-ttu-id="af664-130">撰写</span><span class="sxs-lookup"><span data-stu-id="af664-130">Compose</span></span><br><span data-ttu-id="af664-131">读取</span><span class="sxs-lookup"><span data-stu-id="af664-131">Read</span></span> | <span data-ttu-id="af664-132">String</span><span class="sxs-lookup"><span data-stu-id="af664-132">String</span></span> | [<span data-ttu-id="af664-133">1.1</span><span class="sxs-lookup"><span data-stu-id="af664-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af664-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="af664-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | <span data-ttu-id="af664-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af664-135">ReadItem</span></span> | <span data-ttu-id="af664-136">撰写</span><span class="sxs-lookup"><span data-stu-id="af664-136">Compose</span></span><br><span data-ttu-id="af664-137">读取</span><span class="sxs-lookup"><span data-stu-id="af664-137">Read</span></span> | <span data-ttu-id="af664-138">String</span><span class="sxs-lookup"><span data-stu-id="af664-138">String</span></span> | [<span data-ttu-id="af664-139">1.1</span><span class="sxs-lookup"><span data-stu-id="af664-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

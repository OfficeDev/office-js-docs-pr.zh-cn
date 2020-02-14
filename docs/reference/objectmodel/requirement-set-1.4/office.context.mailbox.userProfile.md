---
title: Office.context.mailbox.userProfile - 要求集 1.4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0532a9971a05412d37334f4c5a4b6b12654f61f3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950990"
---
# <a name="userprofile"></a><span data-ttu-id="1538a-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="1538a-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="1538a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="1538a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="1538a-104">提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="1538a-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1538a-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="1538a-105">Requirements</span></span>

|<span data-ttu-id="1538a-106">要求</span><span class="sxs-lookup"><span data-stu-id="1538a-106">Requirement</span></span>| <span data-ttu-id="1538a-107">值</span><span class="sxs-lookup"><span data-stu-id="1538a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1538a-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1538a-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1538a-109">1.1</span><span class="sxs-lookup"><span data-stu-id="1538a-109">1.1</span></span>|
|[<span data-ttu-id="1538a-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1538a-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1538a-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1538a-111">ReadItem</span></span>|
|[<span data-ttu-id="1538a-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1538a-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1538a-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1538a-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="1538a-114">属性</span><span class="sxs-lookup"><span data-stu-id="1538a-114">Properties</span></span>

| <span data-ttu-id="1538a-115">属性</span><span class="sxs-lookup"><span data-stu-id="1538a-115">Property</span></span> | <span data-ttu-id="1538a-116">最低</span><span class="sxs-lookup"><span data-stu-id="1538a-116">Minimum</span></span><br><span data-ttu-id="1538a-117">权限级别</span><span class="sxs-lookup"><span data-stu-id="1538a-117">permission level</span></span> | <span data-ttu-id="1538a-118">型号</span><span class="sxs-lookup"><span data-stu-id="1538a-118">Modes</span></span> | <span data-ttu-id="1538a-119">返回类型</span><span class="sxs-lookup"><span data-stu-id="1538a-119">Return type</span></span> | <span data-ttu-id="1538a-120">最低</span><span class="sxs-lookup"><span data-stu-id="1538a-120">Minimum</span></span><br><span data-ttu-id="1538a-121">要求集</span><span class="sxs-lookup"><span data-stu-id="1538a-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="1538a-122">displayName</span><span class="sxs-lookup"><span data-stu-id="1538a-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | <span data-ttu-id="1538a-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1538a-123">ReadItem</span></span> | <span data-ttu-id="1538a-124">撰写</span><span class="sxs-lookup"><span data-stu-id="1538a-124">Compose</span></span><br><span data-ttu-id="1538a-125">读取</span><span class="sxs-lookup"><span data-stu-id="1538a-125">Read</span></span> | <span data-ttu-id="1538a-126">字符串</span><span class="sxs-lookup"><span data-stu-id="1538a-126">String</span></span> | [<span data-ttu-id="1538a-127">1.1</span><span class="sxs-lookup"><span data-stu-id="1538a-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1538a-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="1538a-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | <span data-ttu-id="1538a-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1538a-129">ReadItem</span></span> | <span data-ttu-id="1538a-130">撰写</span><span class="sxs-lookup"><span data-stu-id="1538a-130">Compose</span></span><br><span data-ttu-id="1538a-131">读取</span><span class="sxs-lookup"><span data-stu-id="1538a-131">Read</span></span> | <span data-ttu-id="1538a-132">字符串</span><span class="sxs-lookup"><span data-stu-id="1538a-132">String</span></span> | [<span data-ttu-id="1538a-133">1.1</span><span class="sxs-lookup"><span data-stu-id="1538a-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1538a-134">时区</span><span class="sxs-lookup"><span data-stu-id="1538a-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | <span data-ttu-id="1538a-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1538a-135">ReadItem</span></span> | <span data-ttu-id="1538a-136">撰写</span><span class="sxs-lookup"><span data-stu-id="1538a-136">Compose</span></span><br><span data-ttu-id="1538a-137">读取</span><span class="sxs-lookup"><span data-stu-id="1538a-137">Read</span></span> | <span data-ttu-id="1538a-138">字符串</span><span class="sxs-lookup"><span data-stu-id="1538a-138">String</span></span> | [<span data-ttu-id="1538a-139">1.1</span><span class="sxs-lookup"><span data-stu-id="1538a-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

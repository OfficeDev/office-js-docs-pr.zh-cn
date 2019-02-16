---
title: Office.context.mailbox.userProfile - 要求集 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 4a6739c9b463e49d41e320094a4c9cb1a32655f4
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067824"
---
# <a name="userprofile"></a><span data-ttu-id="d8aa1-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="d8aa1-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="d8aa1-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="d8aa1-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8aa1-104">要求</span><span class="sxs-lookup"><span data-stu-id="d8aa1-104">Requirements</span></span>

|<span data-ttu-id="d8aa1-105">要求</span><span class="sxs-lookup"><span data-stu-id="d8aa1-105">Requirement</span></span>| <span data-ttu-id="d8aa1-106">值</span><span class="sxs-lookup"><span data-stu-id="d8aa1-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8aa1-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8aa1-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d8aa1-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d8aa1-108">1.0</span></span>|
|[<span data-ttu-id="d8aa1-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8aa1-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d8aa1-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8aa1-110">ReadItem</span></span>|
|[<span data-ttu-id="d8aa1-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8aa1-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d8aa1-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8aa1-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="d8aa1-113">成员</span><span class="sxs-lookup"><span data-stu-id="d8aa1-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="d8aa1-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="d8aa1-114">displayName :String</span></span>

<span data-ttu-id="d8aa1-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="d8aa1-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="d8aa1-116">Type</span><span class="sxs-lookup"><span data-stu-id="d8aa1-116">Type</span></span>

*   <span data-ttu-id="d8aa1-117">String</span><span class="sxs-lookup"><span data-stu-id="d8aa1-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8aa1-118">要求</span><span class="sxs-lookup"><span data-stu-id="d8aa1-118">Requirements</span></span>

|<span data-ttu-id="d8aa1-119">要求</span><span class="sxs-lookup"><span data-stu-id="d8aa1-119">Requirement</span></span>| <span data-ttu-id="d8aa1-120">值</span><span class="sxs-lookup"><span data-stu-id="d8aa1-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8aa1-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8aa1-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d8aa1-122">1.0</span><span class="sxs-lookup"><span data-stu-id="d8aa1-122">1.0</span></span>|
|[<span data-ttu-id="d8aa1-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8aa1-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d8aa1-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8aa1-124">ReadItem</span></span>|
|[<span data-ttu-id="d8aa1-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8aa1-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d8aa1-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8aa1-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8aa1-127">示例</span><span class="sxs-lookup"><span data-stu-id="d8aa1-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="d8aa1-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="d8aa1-128">emailAddress :String</span></span>

<span data-ttu-id="d8aa1-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="d8aa1-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="d8aa1-130">Type</span><span class="sxs-lookup"><span data-stu-id="d8aa1-130">Type</span></span>

*   <span data-ttu-id="d8aa1-131">String</span><span class="sxs-lookup"><span data-stu-id="d8aa1-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8aa1-132">要求</span><span class="sxs-lookup"><span data-stu-id="d8aa1-132">Requirements</span></span>

|<span data-ttu-id="d8aa1-133">要求</span><span class="sxs-lookup"><span data-stu-id="d8aa1-133">Requirement</span></span>| <span data-ttu-id="d8aa1-134">值</span><span class="sxs-lookup"><span data-stu-id="d8aa1-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8aa1-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8aa1-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d8aa1-136">1.0</span><span class="sxs-lookup"><span data-stu-id="d8aa1-136">1.0</span></span>|
|[<span data-ttu-id="d8aa1-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8aa1-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d8aa1-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8aa1-138">ReadItem</span></span>|
|[<span data-ttu-id="d8aa1-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8aa1-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d8aa1-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8aa1-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8aa1-141">示例</span><span class="sxs-lookup"><span data-stu-id="d8aa1-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="d8aa1-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="d8aa1-142">timeZone :String</span></span>

<span data-ttu-id="d8aa1-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="d8aa1-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="d8aa1-144">Type</span><span class="sxs-lookup"><span data-stu-id="d8aa1-144">Type</span></span>

*   <span data-ttu-id="d8aa1-145">String</span><span class="sxs-lookup"><span data-stu-id="d8aa1-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8aa1-146">要求</span><span class="sxs-lookup"><span data-stu-id="d8aa1-146">Requirements</span></span>

|<span data-ttu-id="d8aa1-147">要求</span><span class="sxs-lookup"><span data-stu-id="d8aa1-147">Requirement</span></span>| <span data-ttu-id="d8aa1-148">值</span><span class="sxs-lookup"><span data-stu-id="d8aa1-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8aa1-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8aa1-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d8aa1-150">1.0</span><span class="sxs-lookup"><span data-stu-id="d8aa1-150">1.0</span></span>|
|[<span data-ttu-id="d8aa1-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8aa1-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d8aa1-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8aa1-152">ReadItem</span></span>|
|[<span data-ttu-id="d8aa1-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8aa1-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d8aa1-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8aa1-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8aa1-155">示例</span><span class="sxs-lookup"><span data-stu-id="d8aa1-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```

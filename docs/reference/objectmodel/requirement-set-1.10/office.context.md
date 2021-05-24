---
title: Office.context - 要求集 1.10
description: Office。适用于使用邮箱 API Outlook集 1.10 的加载项的上下文对象成员。
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: cb189dc3b7b51357dee8ac83bc61795b3ec47ae5
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592030"
---
# <a name="context-mailbox-requirement-set-110"></a><span data-ttu-id="b197e-103">context (Mailbox requirement set 1.10) </span><span class="sxs-lookup"><span data-stu-id="b197e-103">context (Mailbox requirement set 1.10)</span></span>

### <a name="officecontext"></a><span data-ttu-id="b197e-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="b197e-104">[Office](office.md).context</span></span>

<span data-ttu-id="b197e-105">Office.context 提供了外接程序在所有应用程序中使用的共享Office接口。</span><span class="sxs-lookup"><span data-stu-id="b197e-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="b197e-106">此列表仅记录外接程序Outlook接口。有关 Office.context 命名空间的完整列表，请参阅通用 API 中的[Office.context 引用](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="b197e-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b197e-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="b197e-107">Requirements</span></span>

|<span data-ttu-id="b197e-108">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-108">Requirement</span></span>| <span data-ttu-id="b197e-109">值</span><span class="sxs-lookup"><span data-stu-id="b197e-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="b197e-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b197e-111">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-111">1.1</span></span>|
|[<span data-ttu-id="b197e-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b197e-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b197e-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="b197e-114">属性</span><span class="sxs-lookup"><span data-stu-id="b197e-114">Properties</span></span>

| <span data-ttu-id="b197e-115">属性</span><span class="sxs-lookup"><span data-stu-id="b197e-115">Property</span></span> | <span data-ttu-id="b197e-116">模式</span><span class="sxs-lookup"><span data-stu-id="b197e-116">Modes</span></span> | <span data-ttu-id="b197e-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="b197e-117">Return type</span></span> | <span data-ttu-id="b197e-118">最小值</span><span class="sxs-lookup"><span data-stu-id="b197e-118">Minimum</span></span><br><span data-ttu-id="b197e-119">要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b197e-120">auth</span><span class="sxs-lookup"><span data-stu-id="b197e-120">auth</span></span>](#auth-auth) | <span data-ttu-id="b197e-121">撰写</span><span class="sxs-lookup"><span data-stu-id="b197e-121">Compose</span></span><br><span data-ttu-id="b197e-122">阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-122">Read</span></span> | [<span data-ttu-id="b197e-123">Auth</span><span class="sxs-lookup"><span data-stu-id="b197e-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="b197e-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="b197e-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="b197e-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="b197e-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="b197e-126">撰写</span><span class="sxs-lookup"><span data-stu-id="b197e-126">Compose</span></span><br><span data-ttu-id="b197e-127">阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-127">Read</span></span> | <span data-ttu-id="b197e-128">字符串</span><span class="sxs-lookup"><span data-stu-id="b197e-128">String</span></span> | [<span data-ttu-id="b197e-129">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b197e-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="b197e-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="b197e-131">撰写</span><span class="sxs-lookup"><span data-stu-id="b197e-131">Compose</span></span><br><span data-ttu-id="b197e-132">阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-132">Read</span></span> | [<span data-ttu-id="b197e-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b197e-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="b197e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b197e-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="b197e-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="b197e-136">撰写</span><span class="sxs-lookup"><span data-stu-id="b197e-136">Compose</span></span><br><span data-ttu-id="b197e-137">阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-137">Read</span></span> | <span data-ttu-id="b197e-138">字符串</span><span class="sxs-lookup"><span data-stu-id="b197e-138">String</span></span> | [<span data-ttu-id="b197e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b197e-140">host</span><span class="sxs-lookup"><span data-stu-id="b197e-140">host</span></span>](#host-hosttype) | <span data-ttu-id="b197e-141">撰写</span><span class="sxs-lookup"><span data-stu-id="b197e-141">Compose</span></span><br><span data-ttu-id="b197e-142">阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-142">Read</span></span> | [<span data-ttu-id="b197e-143">HostType</span><span class="sxs-lookup"><span data-stu-id="b197e-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="b197e-144">1.5</span><span class="sxs-lookup"><span data-stu-id="b197e-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b197e-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="b197e-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="b197e-146">撰写</span><span class="sxs-lookup"><span data-stu-id="b197e-146">Compose</span></span><br><span data-ttu-id="b197e-147">阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-147">Read</span></span> | [<span data-ttu-id="b197e-148">邮箱</span><span class="sxs-lookup"><span data-stu-id="b197e-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="b197e-149">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b197e-150">平台</span><span class="sxs-lookup"><span data-stu-id="b197e-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="b197e-151">撰写</span><span class="sxs-lookup"><span data-stu-id="b197e-151">Compose</span></span><br><span data-ttu-id="b197e-152">阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-152">Read</span></span> | [<span data-ttu-id="b197e-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b197e-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="b197e-154">1.5</span><span class="sxs-lookup"><span data-stu-id="b197e-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b197e-155">requirements</span><span class="sxs-lookup"><span data-stu-id="b197e-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="b197e-156">撰写</span><span class="sxs-lookup"><span data-stu-id="b197e-156">Compose</span></span><br><span data-ttu-id="b197e-157">阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-157">Read</span></span> | [<span data-ttu-id="b197e-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b197e-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="b197e-159">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b197e-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="b197e-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="b197e-161">撰写</span><span class="sxs-lookup"><span data-stu-id="b197e-161">Compose</span></span><br><span data-ttu-id="b197e-162">阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-162">Read</span></span> | [<span data-ttu-id="b197e-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b197e-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="b197e-164">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b197e-165">ui</span><span class="sxs-lookup"><span data-stu-id="b197e-165">ui</span></span>](#ui-ui) | <span data-ttu-id="b197e-166">撰写</span><span class="sxs-lookup"><span data-stu-id="b197e-166">Compose</span></span><br><span data-ttu-id="b197e-167">阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-167">Read</span></span> | [<span data-ttu-id="b197e-168">UI</span><span class="sxs-lookup"><span data-stu-id="b197e-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="b197e-169">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="b197e-170">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="b197e-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="b197e-171">身份验证 [：Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="b197e-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="b197e-172">通过提供允许 Office 应用程序获取对外接程序 Web 应用程序的访问令牌的方法 ([SSO](../../../outlook/authenticate-a-user-with-an-sso-token.md)) 支持单一登录。</span><span class="sxs-lookup"><span data-stu-id="b197e-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="b197e-173">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="b197e-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="b197e-174">类型</span><span class="sxs-lookup"><span data-stu-id="b197e-174">Type</span></span>

*   [<span data-ttu-id="b197e-175">Auth</span><span class="sxs-lookup"><span data-stu-id="b197e-175">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="b197e-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="b197e-176">Requirements</span></span>

|<span data-ttu-id="b197e-177">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-177">Requirement</span></span>| <span data-ttu-id="b197e-178">值</span><span class="sxs-lookup"><span data-stu-id="b197e-178">Value</span></span>|
|---|---|
|[<span data-ttu-id="b197e-179">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-179">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b197e-180">1.10</span><span class="sxs-lookup"><span data-stu-id="b197e-180">1.10</span></span>|
|[<span data-ttu-id="b197e-181">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b197e-181">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b197e-182">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-182">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b197e-183">示例</span><span class="sxs-lookup"><span data-stu-id="b197e-183">Example</span></span>

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a><span data-ttu-id="b197e-184">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="b197e-184">contentLanguage: String</span></span>

<span data-ttu-id="b197e-185">获取用户 (编辑) 的语言区域设置。</span><span class="sxs-lookup"><span data-stu-id="b197e-185">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="b197e-186">该值 `contentLanguage` 反映当前在客户端 **应用程序中** 由 File **> Options > Language** 指定的Office设置。</span><span class="sxs-lookup"><span data-stu-id="b197e-186">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="b197e-187">类型</span><span class="sxs-lookup"><span data-stu-id="b197e-187">Type</span></span>

*   <span data-ttu-id="b197e-188">String</span><span class="sxs-lookup"><span data-stu-id="b197e-188">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b197e-189">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-189">Requirements</span></span>

|<span data-ttu-id="b197e-190">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-190">Requirement</span></span>| <span data-ttu-id="b197e-191">值</span><span class="sxs-lookup"><span data-stu-id="b197e-191">Value</span></span>|
|---|---|
|[<span data-ttu-id="b197e-192">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-192">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b197e-193">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-193">1.1</span></span>|
|[<span data-ttu-id="b197e-194">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b197e-194">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b197e-195">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-195">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b197e-196">示例</span><span class="sxs-lookup"><span data-stu-id="b197e-196">Example</span></span>

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="b197e-197">diagnostics： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b197e-197">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="b197e-198">获取加载项运行环境的信息。</span><span class="sxs-lookup"><span data-stu-id="b197e-198">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b197e-199">类型</span><span class="sxs-lookup"><span data-stu-id="b197e-199">Type</span></span>

*   [<span data-ttu-id="b197e-200">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b197e-200">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="b197e-201">Requirements</span><span class="sxs-lookup"><span data-stu-id="b197e-201">Requirements</span></span>

|<span data-ttu-id="b197e-202">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-202">Requirement</span></span>| <span data-ttu-id="b197e-203">值</span><span class="sxs-lookup"><span data-stu-id="b197e-203">Value</span></span>|
|---|---|
|[<span data-ttu-id="b197e-204">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-204">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b197e-205">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-205">1.1</span></span>|
|[<span data-ttu-id="b197e-206">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b197e-206">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b197e-207">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-207">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b197e-208">示例</span><span class="sxs-lookup"><span data-stu-id="b197e-208">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="b197e-209">displayLanguage：String</span><span class="sxs-lookup"><span data-stu-id="b197e-209">displayLanguage: String</span></span>

<span data-ttu-id="b197e-210">获取区域设置 (语言) RFC 1766 语言标记格式，该标记格式由用户为 Office 客户端应用程序的 UI 指定。</span><span class="sxs-lookup"><span data-stu-id="b197e-210">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="b197e-211">该值反映当前显示语言设置，该设置由 > `displayLanguage` **客户端** 应用程序中>选项Office语言。 </span><span class="sxs-lookup"><span data-stu-id="b197e-211">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="b197e-212">类型</span><span class="sxs-lookup"><span data-stu-id="b197e-212">Type</span></span>

*   <span data-ttu-id="b197e-213">String</span><span class="sxs-lookup"><span data-stu-id="b197e-213">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b197e-214">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-214">Requirements</span></span>

|<span data-ttu-id="b197e-215">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-215">Requirement</span></span>| <span data-ttu-id="b197e-216">值</span><span class="sxs-lookup"><span data-stu-id="b197e-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="b197e-217">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-217">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b197e-218">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-218">1.1</span></span>|
|[<span data-ttu-id="b197e-219">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b197e-219">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b197e-220">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b197e-221">示例</span><span class="sxs-lookup"><span data-stu-id="b197e-221">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="host-hosttype"></a><span data-ttu-id="b197e-222">host： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="b197e-222">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="b197e-223">获取Office加载项的加载项应用程序。</span><span class="sxs-lookup"><span data-stu-id="b197e-223">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b197e-224">或者，您可以使用[Office.context.diagnostics](#diagnostics-contextinformation)属性获取主机。</span><span class="sxs-lookup"><span data-stu-id="b197e-224">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="b197e-225">类型</span><span class="sxs-lookup"><span data-stu-id="b197e-225">Type</span></span>

*   [<span data-ttu-id="b197e-226">HostType</span><span class="sxs-lookup"><span data-stu-id="b197e-226">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="b197e-227">Requirements</span><span class="sxs-lookup"><span data-stu-id="b197e-227">Requirements</span></span>

|<span data-ttu-id="b197e-228">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-228">Requirement</span></span>| <span data-ttu-id="b197e-229">值</span><span class="sxs-lookup"><span data-stu-id="b197e-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="b197e-230">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-230">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b197e-231">1.5</span><span class="sxs-lookup"><span data-stu-id="b197e-231">1.5</span></span>|
|[<span data-ttu-id="b197e-232">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b197e-232">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b197e-233">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b197e-234">示例</span><span class="sxs-lookup"><span data-stu-id="b197e-234">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="b197e-235">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="b197e-235">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="b197e-236">提供运行加载项的平台。</span><span class="sxs-lookup"><span data-stu-id="b197e-236">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="b197e-237">或者，您可以使用[Office.context.diagnostics](#diagnostics-contextinformation)属性获取平台。</span><span class="sxs-lookup"><span data-stu-id="b197e-237">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b197e-238">类型</span><span class="sxs-lookup"><span data-stu-id="b197e-238">Type</span></span>

*   [<span data-ttu-id="b197e-239">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b197e-239">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="b197e-240">Requirements</span><span class="sxs-lookup"><span data-stu-id="b197e-240">Requirements</span></span>

|<span data-ttu-id="b197e-241">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-241">Requirement</span></span>| <span data-ttu-id="b197e-242">值</span><span class="sxs-lookup"><span data-stu-id="b197e-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="b197e-243">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b197e-244">1.5</span><span class="sxs-lookup"><span data-stu-id="b197e-244">1.5</span></span>|
|[<span data-ttu-id="b197e-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b197e-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b197e-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b197e-247">示例</span><span class="sxs-lookup"><span data-stu-id="b197e-247">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="b197e-248">requirements： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="b197e-248">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="b197e-249">提供用于确定当前应用程序和平台上支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="b197e-249">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b197e-250">类型</span><span class="sxs-lookup"><span data-stu-id="b197e-250">Type</span></span>

*   [<span data-ttu-id="b197e-251">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b197e-251">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="b197e-252">Requirements</span><span class="sxs-lookup"><span data-stu-id="b197e-252">Requirements</span></span>

|<span data-ttu-id="b197e-253">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-253">Requirement</span></span>| <span data-ttu-id="b197e-254">值</span><span class="sxs-lookup"><span data-stu-id="b197e-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="b197e-255">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-255">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b197e-256">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-256">1.1</span></span>|
|[<span data-ttu-id="b197e-257">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b197e-257">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b197e-258">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b197e-259">示例</span><span class="sxs-lookup"><span data-stu-id="b197e-259">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="b197e-260">[roamingSettings：RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="b197e-260">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="b197e-261">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="b197e-261">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b197e-262">该对象允许您存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时可供该外接程序使用 `RoamingSettings` 。</span><span class="sxs-lookup"><span data-stu-id="b197e-262">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b197e-263">类型</span><span class="sxs-lookup"><span data-stu-id="b197e-263">Type</span></span>

*   [<span data-ttu-id="b197e-264">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b197e-264">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="b197e-265">Requirements</span><span class="sxs-lookup"><span data-stu-id="b197e-265">Requirements</span></span>

|<span data-ttu-id="b197e-266">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-266">Requirement</span></span>| <span data-ttu-id="b197e-267">值</span><span class="sxs-lookup"><span data-stu-id="b197e-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="b197e-268">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-268">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b197e-269">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-269">1.1</span></span>|
|[<span data-ttu-id="b197e-270">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b197e-270">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="b197e-271">受限</span><span class="sxs-lookup"><span data-stu-id="b197e-271">Restricted</span></span>|
|[<span data-ttu-id="b197e-272">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b197e-272">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b197e-273">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="b197e-274">[ui：UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="b197e-274">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="b197e-275">提供可用于在加载项中创建和操作 UI 组件（如对话框）Office方法。</span><span class="sxs-lookup"><span data-stu-id="b197e-275">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b197e-276">类型</span><span class="sxs-lookup"><span data-stu-id="b197e-276">Type</span></span>

*   [<span data-ttu-id="b197e-277">UI</span><span class="sxs-lookup"><span data-stu-id="b197e-277">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="b197e-278">Requirements</span><span class="sxs-lookup"><span data-stu-id="b197e-278">Requirements</span></span>

|<span data-ttu-id="b197e-279">要求</span><span class="sxs-lookup"><span data-stu-id="b197e-279">Requirement</span></span>| <span data-ttu-id="b197e-280">值</span><span class="sxs-lookup"><span data-stu-id="b197e-280">Value</span></span>|
|---|---|
|[<span data-ttu-id="b197e-281">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b197e-281">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b197e-282">1.1</span><span class="sxs-lookup"><span data-stu-id="b197e-282">1.1</span></span>|
|[<span data-ttu-id="b197e-283">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b197e-283">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b197e-284">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b197e-284">Compose or Read</span></span>|

---
title: Office。上下文要求集1。9
description: 使用邮箱 API 要求集1.9 的 Outlook 外接程序可用的 Office 对象成员。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 6b2657d1e608bd1820d3814d9a6bfab67681824c
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628054"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="81549-103"> (邮箱要求集1.9 的上下文) </span><span class="sxs-lookup"><span data-stu-id="81549-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="81549-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="81549-104">[Office](office.md).context</span></span>

<span data-ttu-id="81549-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="81549-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="81549-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅 [通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true)"。</span><span class="sxs-lookup"><span data-stu-id="81549-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="81549-107">要求</span><span class="sxs-lookup"><span data-stu-id="81549-107">Requirements</span></span>

|<span data-ttu-id="81549-108">要求</span><span class="sxs-lookup"><span data-stu-id="81549-108">Requirement</span></span>| <span data-ttu-id="81549-109">值</span><span class="sxs-lookup"><span data-stu-id="81549-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="81549-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="81549-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="81549-111">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-111">1.1</span></span>|
|[<span data-ttu-id="81549-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="81549-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="81549-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="81549-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="81549-114">属性</span><span class="sxs-lookup"><span data-stu-id="81549-114">Properties</span></span>

| <span data-ttu-id="81549-115">属性</span><span class="sxs-lookup"><span data-stu-id="81549-115">Property</span></span> | <span data-ttu-id="81549-116">型号</span><span class="sxs-lookup"><span data-stu-id="81549-116">Modes</span></span> | <span data-ttu-id="81549-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="81549-117">Return type</span></span> | <span data-ttu-id="81549-118">最小值</span><span class="sxs-lookup"><span data-stu-id="81549-118">Minimum</span></span><br><span data-ttu-id="81549-119">要求集</span><span class="sxs-lookup"><span data-stu-id="81549-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="81549-120">认证</span><span class="sxs-lookup"><span data-stu-id="81549-120">auth</span></span>](#auth-auth) | <span data-ttu-id="81549-121">撰写</span><span class="sxs-lookup"><span data-stu-id="81549-121">Compose</span></span><br><span data-ttu-id="81549-122">读取</span><span class="sxs-lookup"><span data-stu-id="81549-122">Read</span></span> | [<span data-ttu-id="81549-123">Auth</span><span class="sxs-lookup"><span data-stu-id="81549-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="81549-124">IdentityAPI 1。3</span><span class="sxs-lookup"><span data-stu-id="81549-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="81549-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="81549-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="81549-126">撰写</span><span class="sxs-lookup"><span data-stu-id="81549-126">Compose</span></span><br><span data-ttu-id="81549-127">读取</span><span class="sxs-lookup"><span data-stu-id="81549-127">Read</span></span> | <span data-ttu-id="81549-128">String</span><span class="sxs-lookup"><span data-stu-id="81549-128">String</span></span> | [<span data-ttu-id="81549-129">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="81549-130">过程</span><span class="sxs-lookup"><span data-stu-id="81549-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="81549-131">撰写</span><span class="sxs-lookup"><span data-stu-id="81549-131">Compose</span></span><br><span data-ttu-id="81549-132">读取</span><span class="sxs-lookup"><span data-stu-id="81549-132">Read</span></span> | [<span data-ttu-id="81549-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="81549-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="81549-134">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="81549-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="81549-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="81549-136">撰写</span><span class="sxs-lookup"><span data-stu-id="81549-136">Compose</span></span><br><span data-ttu-id="81549-137">读取</span><span class="sxs-lookup"><span data-stu-id="81549-137">Read</span></span> | <span data-ttu-id="81549-138">String</span><span class="sxs-lookup"><span data-stu-id="81549-138">String</span></span> | [<span data-ttu-id="81549-139">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="81549-140">host</span><span class="sxs-lookup"><span data-stu-id="81549-140">host</span></span>](#host-hosttype) | <span data-ttu-id="81549-141">撰写</span><span class="sxs-lookup"><span data-stu-id="81549-141">Compose</span></span><br><span data-ttu-id="81549-142">读取</span><span class="sxs-lookup"><span data-stu-id="81549-142">Read</span></span> | [<span data-ttu-id="81549-143">HostType</span><span class="sxs-lookup"><span data-stu-id="81549-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="81549-144">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="81549-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="81549-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="81549-146">撰写</span><span class="sxs-lookup"><span data-stu-id="81549-146">Compose</span></span><br><span data-ttu-id="81549-147">读取</span><span class="sxs-lookup"><span data-stu-id="81549-147">Read</span></span> | [<span data-ttu-id="81549-148">邮箱</span><span class="sxs-lookup"><span data-stu-id="81549-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="81549-149">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="81549-150">平台</span><span class="sxs-lookup"><span data-stu-id="81549-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="81549-151">撰写</span><span class="sxs-lookup"><span data-stu-id="81549-151">Compose</span></span><br><span data-ttu-id="81549-152">读取</span><span class="sxs-lookup"><span data-stu-id="81549-152">Read</span></span> | [<span data-ttu-id="81549-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="81549-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="81549-154">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="81549-155">满足</span><span class="sxs-lookup"><span data-stu-id="81549-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="81549-156">撰写</span><span class="sxs-lookup"><span data-stu-id="81549-156">Compose</span></span><br><span data-ttu-id="81549-157">读取</span><span class="sxs-lookup"><span data-stu-id="81549-157">Read</span></span> | [<span data-ttu-id="81549-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="81549-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="81549-159">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="81549-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="81549-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="81549-161">撰写</span><span class="sxs-lookup"><span data-stu-id="81549-161">Compose</span></span><br><span data-ttu-id="81549-162">读取</span><span class="sxs-lookup"><span data-stu-id="81549-162">Read</span></span> | [<span data-ttu-id="81549-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="81549-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="81549-164">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="81549-165">ui</span><span class="sxs-lookup"><span data-stu-id="81549-165">ui</span></span>](#ui-ui) | <span data-ttu-id="81549-166">撰写</span><span class="sxs-lookup"><span data-stu-id="81549-166">Compose</span></span><br><span data-ttu-id="81549-167">读取</span><span class="sxs-lookup"><span data-stu-id="81549-167">Read</span></span> | [<span data-ttu-id="81549-168">UI</span><span class="sxs-lookup"><span data-stu-id="81549-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="81549-169">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="81549-170">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="81549-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="81549-171">auth： [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="81549-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="81549-172">通过提供允许 Office 应用程序获取对加载项 web 应用程序的访问令牌的方法，支持 [单一登录 (SSO) ](../../../outlook/authenticate-a-user-with-an-sso-token.md) 。</span><span class="sxs-lookup"><span data-stu-id="81549-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="81549-173">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="81549-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="81549-174">请参阅 [IdentityAPI 1.3 要求集](../../requirement-sets/identity-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="81549-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="81549-175">类型</span><span class="sxs-lookup"><span data-stu-id="81549-175">Type</span></span>

*   [<span data-ttu-id="81549-176">Auth</span><span class="sxs-lookup"><span data-stu-id="81549-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="81549-177">要求</span><span class="sxs-lookup"><span data-stu-id="81549-177">Requirements</span></span>

|<span data-ttu-id="81549-178">要求</span><span class="sxs-lookup"><span data-stu-id="81549-178">Requirement</span></span>| <span data-ttu-id="81549-179">值</span><span class="sxs-lookup"><span data-stu-id="81549-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="81549-180">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="81549-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="81549-181">无</span><span class="sxs-lookup"><span data-stu-id="81549-181">N/A</span></span>|
|[<span data-ttu-id="81549-182">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="81549-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="81549-183">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="81549-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="81549-184">示例</span><span class="sxs-lookup"><span data-stu-id="81549-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="81549-185">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="81549-185">contentLanguage: String</span></span>

<span data-ttu-id="81549-186">获取用户指定的用于编辑项目的区域设置 (语言) 。</span><span class="sxs-lookup"><span data-stu-id="81549-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="81549-187">此 `contentLanguage` 值反映了使用 Office 客户端应用程序中的 "**文件 > 选项" > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="81549-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="81549-188">类型</span><span class="sxs-lookup"><span data-stu-id="81549-188">Type</span></span>

*   <span data-ttu-id="81549-189">String</span><span class="sxs-lookup"><span data-stu-id="81549-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="81549-190">要求</span><span class="sxs-lookup"><span data-stu-id="81549-190">Requirements</span></span>

|<span data-ttu-id="81549-191">要求</span><span class="sxs-lookup"><span data-stu-id="81549-191">Requirement</span></span>| <span data-ttu-id="81549-192">值</span><span class="sxs-lookup"><span data-stu-id="81549-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="81549-193">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="81549-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="81549-194">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-194">1.1</span></span>|
|[<span data-ttu-id="81549-195">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="81549-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="81549-196">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="81549-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="81549-197">示例</span><span class="sxs-lookup"><span data-stu-id="81549-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="81549-198">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="81549-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="81549-199">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="81549-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="81549-200">类型</span><span class="sxs-lookup"><span data-stu-id="81549-200">Type</span></span>

*   [<span data-ttu-id="81549-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="81549-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="81549-202">要求</span><span class="sxs-lookup"><span data-stu-id="81549-202">Requirements</span></span>

|<span data-ttu-id="81549-203">要求</span><span class="sxs-lookup"><span data-stu-id="81549-203">Requirement</span></span>| <span data-ttu-id="81549-204">值</span><span class="sxs-lookup"><span data-stu-id="81549-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="81549-205">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="81549-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="81549-206">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-206">1.1</span></span>|
|[<span data-ttu-id="81549-207">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="81549-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="81549-208">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="81549-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="81549-209">示例</span><span class="sxs-lookup"><span data-stu-id="81549-209">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="81549-210">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="81549-210">displayLanguage: String</span></span>

<span data-ttu-id="81549-211">获取用户为 Office 客户端应用程序的 UI 指定的 RFC 1766 语言标记格式中 (语言) 的区域设置。</span><span class="sxs-lookup"><span data-stu-id="81549-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="81549-212">此 `displayLanguage` 值反映了使用 Office 客户端应用程序中的**File > Options > 语言**指定的当前**显示语言**设置。</span><span class="sxs-lookup"><span data-stu-id="81549-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="81549-213">类型</span><span class="sxs-lookup"><span data-stu-id="81549-213">Type</span></span>

*   <span data-ttu-id="81549-214">String</span><span class="sxs-lookup"><span data-stu-id="81549-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="81549-215">要求</span><span class="sxs-lookup"><span data-stu-id="81549-215">Requirements</span></span>

|<span data-ttu-id="81549-216">要求</span><span class="sxs-lookup"><span data-stu-id="81549-216">Requirement</span></span>| <span data-ttu-id="81549-217">值</span><span class="sxs-lookup"><span data-stu-id="81549-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="81549-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="81549-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="81549-219">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-219">1.1</span></span>|
|[<span data-ttu-id="81549-220">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="81549-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="81549-221">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="81549-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="81549-222">示例</span><span class="sxs-lookup"><span data-stu-id="81549-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="81549-223">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="81549-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="81549-224">获取承载外接程序的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="81549-224">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="81549-225">类型</span><span class="sxs-lookup"><span data-stu-id="81549-225">Type</span></span>

*   [<span data-ttu-id="81549-226">HostType</span><span class="sxs-lookup"><span data-stu-id="81549-226">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="81549-227">要求</span><span class="sxs-lookup"><span data-stu-id="81549-227">Requirements</span></span>

|<span data-ttu-id="81549-228">要求</span><span class="sxs-lookup"><span data-stu-id="81549-228">Requirement</span></span>| <span data-ttu-id="81549-229">值</span><span class="sxs-lookup"><span data-stu-id="81549-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="81549-230">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="81549-230">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="81549-231">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-231">1.1</span></span>|
|[<span data-ttu-id="81549-232">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="81549-232">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="81549-233">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="81549-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="81549-234">示例</span><span class="sxs-lookup"><span data-stu-id="81549-234">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="81549-235">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="81549-235">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="81549-236">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="81549-236">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="81549-237">类型</span><span class="sxs-lookup"><span data-stu-id="81549-237">Type</span></span>

*   [<span data-ttu-id="81549-238">PlatformType</span><span class="sxs-lookup"><span data-stu-id="81549-238">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="81549-239">要求</span><span class="sxs-lookup"><span data-stu-id="81549-239">Requirements</span></span>

|<span data-ttu-id="81549-240">要求</span><span class="sxs-lookup"><span data-stu-id="81549-240">Requirement</span></span>| <span data-ttu-id="81549-241">值</span><span class="sxs-lookup"><span data-stu-id="81549-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="81549-242">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="81549-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="81549-243">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-243">1.1</span></span>|
|[<span data-ttu-id="81549-244">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="81549-244">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="81549-245">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="81549-245">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="81549-246">示例</span><span class="sxs-lookup"><span data-stu-id="81549-246">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="81549-247">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="81549-247">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="81549-248">提供用于确定当前应用程序和平台支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="81549-248">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="81549-249">类型</span><span class="sxs-lookup"><span data-stu-id="81549-249">Type</span></span>

*   [<span data-ttu-id="81549-250">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="81549-250">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="81549-251">要求</span><span class="sxs-lookup"><span data-stu-id="81549-251">Requirements</span></span>

|<span data-ttu-id="81549-252">要求</span><span class="sxs-lookup"><span data-stu-id="81549-252">Requirement</span></span>| <span data-ttu-id="81549-253">值</span><span class="sxs-lookup"><span data-stu-id="81549-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="81549-254">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="81549-254">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="81549-255">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-255">1.1</span></span>|
|[<span data-ttu-id="81549-256">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="81549-256">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="81549-257">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="81549-257">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="81549-258">示例</span><span class="sxs-lookup"><span data-stu-id="81549-258">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="81549-259">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="81549-259">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="81549-260">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="81549-260">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="81549-261">该 `RoamingSettings` 对象使您可以存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时都可使用该加载项。</span><span class="sxs-lookup"><span data-stu-id="81549-261">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="81549-262">类型</span><span class="sxs-lookup"><span data-stu-id="81549-262">Type</span></span>

*   [<span data-ttu-id="81549-263">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="81549-263">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="81549-264">要求</span><span class="sxs-lookup"><span data-stu-id="81549-264">Requirements</span></span>

|<span data-ttu-id="81549-265">要求</span><span class="sxs-lookup"><span data-stu-id="81549-265">Requirement</span></span>| <span data-ttu-id="81549-266">值</span><span class="sxs-lookup"><span data-stu-id="81549-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="81549-267">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="81549-267">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="81549-268">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-268">1.1</span></span>|
|[<span data-ttu-id="81549-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="81549-269">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="81549-270">受限</span><span class="sxs-lookup"><span data-stu-id="81549-270">Restricted</span></span>|
|[<span data-ttu-id="81549-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="81549-271">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="81549-272">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="81549-272">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="81549-273">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="81549-273">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="81549-274">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="81549-274">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="81549-275">类型</span><span class="sxs-lookup"><span data-stu-id="81549-275">Type</span></span>

*   [<span data-ttu-id="81549-276">UI</span><span class="sxs-lookup"><span data-stu-id="81549-276">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="81549-277">要求</span><span class="sxs-lookup"><span data-stu-id="81549-277">Requirements</span></span>

|<span data-ttu-id="81549-278">要求</span><span class="sxs-lookup"><span data-stu-id="81549-278">Requirement</span></span>| <span data-ttu-id="81549-279">值</span><span class="sxs-lookup"><span data-stu-id="81549-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="81549-280">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="81549-280">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="81549-281">1.1</span><span class="sxs-lookup"><span data-stu-id="81549-281">1.1</span></span>|
|[<span data-ttu-id="81549-282">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="81549-282">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="81549-283">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="81549-283">Compose or Read</span></span>|

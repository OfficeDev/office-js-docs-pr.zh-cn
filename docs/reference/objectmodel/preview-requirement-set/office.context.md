---
title: Office. context-预览要求集
description: 使用邮箱 API preview 要求集的 Outlook 外接程序可用的 Context 对象成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 0e0ea973032bb5cd854856920e192522f90a26a1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612022"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="d91c8-103">context （邮箱预览要求集）</span><span class="sxs-lookup"><span data-stu-id="d91c8-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="d91c8-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="d91c8-104">[Office](office.md).context</span></span>

<span data-ttu-id="d91c8-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="d91c8-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="d91c8-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-preview)"。</span><span class="sxs-lookup"><span data-stu-id="d91c8-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d91c8-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="d91c8-107">Requirements</span></span>

|<span data-ttu-id="d91c8-108">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-108">Requirement</span></span>| <span data-ttu-id="d91c8-109">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-111">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-111">1.1</span></span>|
|[<span data-ttu-id="d91c8-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="d91c8-114">属性</span><span class="sxs-lookup"><span data-stu-id="d91c8-114">Properties</span></span>

| <span data-ttu-id="d91c8-115">属性</span><span class="sxs-lookup"><span data-stu-id="d91c8-115">Property</span></span> | <span data-ttu-id="d91c8-116">型号</span><span class="sxs-lookup"><span data-stu-id="d91c8-116">Modes</span></span> | <span data-ttu-id="d91c8-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-117">Return type</span></span> | <span data-ttu-id="d91c8-118">最低</span><span class="sxs-lookup"><span data-stu-id="d91c8-118">Minimum</span></span><br><span data-ttu-id="d91c8-119">要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d91c8-120">认证</span><span class="sxs-lookup"><span data-stu-id="d91c8-120">auth</span></span>](#auth-auth) | <span data-ttu-id="d91c8-121">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-121">Compose</span></span><br><span data-ttu-id="d91c8-122">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-122">Read</span></span> | [<span data-ttu-id="d91c8-123">Auth</span><span class="sxs-lookup"><span data-stu-id="d91c8-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="d91c8-124">预览</span><span class="sxs-lookup"><span data-stu-id="d91c8-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="d91c8-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="d91c8-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="d91c8-126">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-126">Compose</span></span><br><span data-ttu-id="d91c8-127">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-127">Read</span></span> | <span data-ttu-id="d91c8-128">String</span><span class="sxs-lookup"><span data-stu-id="d91c8-128">String</span></span> | [<span data-ttu-id="d91c8-129">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d91c8-130">过程</span><span class="sxs-lookup"><span data-stu-id="d91c8-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="d91c8-131">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-131">Compose</span></span><br><span data-ttu-id="d91c8-132">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-132">Read</span></span> | [<span data-ttu-id="d91c8-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="d91c8-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="d91c8-134">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d91c8-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="d91c8-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="d91c8-136">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-136">Compose</span></span><br><span data-ttu-id="d91c8-137">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-137">Read</span></span> | <span data-ttu-id="d91c8-138">String</span><span class="sxs-lookup"><span data-stu-id="d91c8-138">String</span></span> | [<span data-ttu-id="d91c8-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d91c8-140">host</span><span class="sxs-lookup"><span data-stu-id="d91c8-140">host</span></span>](#host-hosttype) | <span data-ttu-id="d91c8-141">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-141">Compose</span></span><br><span data-ttu-id="d91c8-142">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-142">Read</span></span> | [<span data-ttu-id="d91c8-143">HostType</span><span class="sxs-lookup"><span data-stu-id="d91c8-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="d91c8-144">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d91c8-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="d91c8-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="d91c8-146">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-146">Compose</span></span><br><span data-ttu-id="d91c8-147">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-147">Read</span></span> | [<span data-ttu-id="d91c8-148">邮箱</span><span class="sxs-lookup"><span data-stu-id="d91c8-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="d91c8-149">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d91c8-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="d91c8-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="d91c8-151">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-151">Compose</span></span><br><span data-ttu-id="d91c8-152">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-152">Read</span></span> | [<span data-ttu-id="d91c8-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="d91c8-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="d91c8-154">预览</span><span class="sxs-lookup"><span data-stu-id="d91c8-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="d91c8-155">平台</span><span class="sxs-lookup"><span data-stu-id="d91c8-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="d91c8-156">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-156">Compose</span></span><br><span data-ttu-id="d91c8-157">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-157">Read</span></span> | [<span data-ttu-id="d91c8-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="d91c8-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="d91c8-159">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d91c8-160">满足</span><span class="sxs-lookup"><span data-stu-id="d91c8-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="d91c8-161">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-161">Compose</span></span><br><span data-ttu-id="d91c8-162">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-162">Read</span></span> | [<span data-ttu-id="d91c8-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="d91c8-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="d91c8-164">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d91c8-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="d91c8-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="d91c8-166">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-166">Compose</span></span><br><span data-ttu-id="d91c8-167">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-167">Read</span></span> | [<span data-ttu-id="d91c8-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d91c8-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="d91c8-169">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d91c8-170">ui</span><span class="sxs-lookup"><span data-stu-id="d91c8-170">ui</span></span>](#ui-ui) | <span data-ttu-id="d91c8-171">撰写</span><span class="sxs-lookup"><span data-stu-id="d91c8-171">Compose</span></span><br><span data-ttu-id="d91c8-172">Read</span><span class="sxs-lookup"><span data-stu-id="d91c8-172">Read</span></span> | [<span data-ttu-id="d91c8-173">UI</span><span class="sxs-lookup"><span data-stu-id="d91c8-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="d91c8-174">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="d91c8-175">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="d91c8-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="d91c8-176">auth： [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="d91c8-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="d91c8-177">通过提供一种方法来支持[单一登录（SSO）](../../../outlook/authenticate-a-user-with-an-sso-token.md) ，使 Office 主机能够获取加载项的 web 应用程序的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="d91c8-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="d91c8-178">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="d91c8-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="d91c8-179">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-179">Type</span></span>

*   [<span data-ttu-id="d91c8-180">Auth</span><span class="sxs-lookup"><span data-stu-id="d91c8-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="d91c8-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="d91c8-181">Requirements</span></span>

|<span data-ttu-id="d91c8-182">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-182">Requirement</span></span>| <span data-ttu-id="d91c8-183">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-185">预览</span><span class="sxs-lookup"><span data-stu-id="d91c8-185">Preview</span></span>|
|[<span data-ttu-id="d91c8-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d91c8-188">示例</span><span class="sxs-lookup"><span data-stu-id="d91c8-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="d91c8-189">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="d91c8-189">contentLanguage: String</span></span>

<span data-ttu-id="d91c8-190">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="d91c8-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="d91c8-191">此 `contentLanguage` 值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="d91c8-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="d91c8-192">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-192">Type</span></span>

*   <span data-ttu-id="d91c8-193">String</span><span class="sxs-lookup"><span data-stu-id="d91c8-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d91c8-194">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-194">Requirements</span></span>

|<span data-ttu-id="d91c8-195">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-195">Requirement</span></span>| <span data-ttu-id="d91c8-196">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-198">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-198">1.1</span></span>|
|[<span data-ttu-id="d91c8-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d91c8-201">示例</span><span class="sxs-lookup"><span data-stu-id="d91c8-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="d91c8-202">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="d91c8-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="d91c8-203">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="d91c8-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="d91c8-204">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-204">Type</span></span>

*   [<span data-ttu-id="d91c8-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="d91c8-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="d91c8-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="d91c8-206">Requirements</span></span>

|<span data-ttu-id="d91c8-207">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-207">Requirement</span></span>| <span data-ttu-id="d91c8-208">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-209">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-210">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-210">1.1</span></span>|
|[<span data-ttu-id="d91c8-211">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-212">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d91c8-213">示例</span><span class="sxs-lookup"><span data-stu-id="d91c8-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="d91c8-214">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="d91c8-214">displayLanguage: String</span></span>

<span data-ttu-id="d91c8-215">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="d91c8-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="d91c8-216">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="d91c8-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="d91c8-217">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-217">Type</span></span>

*   <span data-ttu-id="d91c8-218">String</span><span class="sxs-lookup"><span data-stu-id="d91c8-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d91c8-219">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-219">Requirements</span></span>

|<span data-ttu-id="d91c8-220">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-220">Requirement</span></span>| <span data-ttu-id="d91c8-221">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-223">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-223">1.1</span></span>|
|[<span data-ttu-id="d91c8-224">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-225">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d91c8-226">示例</span><span class="sxs-lookup"><span data-stu-id="d91c8-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="d91c8-227">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="d91c8-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="d91c8-228">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="d91c8-228">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="d91c8-229">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-229">Type</span></span>

*   [<span data-ttu-id="d91c8-230">HostType</span><span class="sxs-lookup"><span data-stu-id="d91c8-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="d91c8-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="d91c8-231">Requirements</span></span>

|<span data-ttu-id="d91c8-232">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-232">Requirement</span></span>| <span data-ttu-id="d91c8-233">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-234">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-235">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-235">1.1</span></span>|
|[<span data-ttu-id="d91c8-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d91c8-238">示例</span><span class="sxs-lookup"><span data-stu-id="d91c8-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="d91c8-239">officeTheme： [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="d91c8-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="d91c8-240">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="d91c8-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="d91c8-241">此成员仅在 Windows 中的 Outlook 中受支持。</span><span class="sxs-lookup"><span data-stu-id="d91c8-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="d91c8-242">使用 Office 主题颜色，您可以将加载项的配色方案与用户选择的当前 Office 主题进行协调，以供用户使用**office > Office 帐户 > Office 主题 UI**，该用户在所有 Office 主机应用程序中应用。</span><span class="sxs-lookup"><span data-stu-id="d91c8-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="d91c8-243">使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="d91c8-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="d91c8-244">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-244">Type</span></span>

*   [<span data-ttu-id="d91c8-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="d91c8-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="d91c8-246">属性：</span><span class="sxs-lookup"><span data-stu-id="d91c8-246">Properties:</span></span>

|<span data-ttu-id="d91c8-247">名称</span><span class="sxs-lookup"><span data-stu-id="d91c8-247">Name</span></span>| <span data-ttu-id="d91c8-248">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-248">Type</span></span>| <span data-ttu-id="d91c8-249">说明</span><span class="sxs-lookup"><span data-stu-id="d91c8-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="d91c8-250">String</span><span class="sxs-lookup"><span data-stu-id="d91c8-250">String</span></span>|<span data-ttu-id="d91c8-251">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="d91c8-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="d91c8-252">String</span><span class="sxs-lookup"><span data-stu-id="d91c8-252">String</span></span>|<span data-ttu-id="d91c8-253">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="d91c8-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="d91c8-254">String</span><span class="sxs-lookup"><span data-stu-id="d91c8-254">String</span></span>|<span data-ttu-id="d91c8-255">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="d91c8-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="d91c8-256">字符串</span><span class="sxs-lookup"><span data-stu-id="d91c8-256">String</span></span>|<span data-ttu-id="d91c8-257">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="d91c8-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d91c8-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="d91c8-258">Requirements</span></span>

|<span data-ttu-id="d91c8-259">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-259">Requirement</span></span>| <span data-ttu-id="d91c8-260">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-262">预览</span><span class="sxs-lookup"><span data-stu-id="d91c8-262">Preview</span></span>|
|[<span data-ttu-id="d91c8-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d91c8-265">示例</span><span class="sxs-lookup"><span data-stu-id="d91c8-265">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="d91c8-266">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="d91c8-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="d91c8-267">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="d91c8-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="d91c8-268">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-268">Type</span></span>

*   [<span data-ttu-id="d91c8-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="d91c8-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="d91c8-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="d91c8-270">Requirements</span></span>

|<span data-ttu-id="d91c8-271">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-271">Requirement</span></span>| <span data-ttu-id="d91c8-272">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-273">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-274">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-274">1.1</span></span>|
|[<span data-ttu-id="d91c8-275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-276">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d91c8-277">示例</span><span class="sxs-lookup"><span data-stu-id="d91c8-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="d91c8-278">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="d91c8-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="d91c8-279">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="d91c8-279">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="d91c8-280">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-280">Type</span></span>

*   [<span data-ttu-id="d91c8-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="d91c8-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="d91c8-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="d91c8-282">Requirements</span></span>

|<span data-ttu-id="d91c8-283">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-283">Requirement</span></span>| <span data-ttu-id="d91c8-284">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-285">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-286">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-286">1.1</span></span>|
|[<span data-ttu-id="d91c8-287">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-288">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d91c8-289">示例</span><span class="sxs-lookup"><span data-stu-id="d91c8-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="d91c8-290">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="d91c8-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="d91c8-291">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="d91c8-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="d91c8-292">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="d91c8-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="d91c8-293">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-293">Type</span></span>

*   [<span data-ttu-id="d91c8-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d91c8-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="d91c8-295">Requirements</span><span class="sxs-lookup"><span data-stu-id="d91c8-295">Requirements</span></span>

|<span data-ttu-id="d91c8-296">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-296">Requirement</span></span>| <span data-ttu-id="d91c8-297">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-298">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-299">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-299">1.1</span></span>|
|[<span data-ttu-id="d91c8-300">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d91c8-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="d91c8-301">受限</span><span class="sxs-lookup"><span data-stu-id="d91c8-301">Restricted</span></span>|
|[<span data-ttu-id="d91c8-302">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-303">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="d91c8-304">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="d91c8-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="d91c8-305">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="d91c8-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="d91c8-306">类型</span><span class="sxs-lookup"><span data-stu-id="d91c8-306">Type</span></span>

*   [<span data-ttu-id="d91c8-307">UI</span><span class="sxs-lookup"><span data-stu-id="d91c8-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="d91c8-308">Requirements</span><span class="sxs-lookup"><span data-stu-id="d91c8-308">Requirements</span></span>

|<span data-ttu-id="d91c8-309">要求</span><span class="sxs-lookup"><span data-stu-id="d91c8-309">Requirement</span></span>| <span data-ttu-id="d91c8-310">值</span><span class="sxs-lookup"><span data-stu-id="d91c8-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="d91c8-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d91c8-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d91c8-312">1.1</span><span class="sxs-lookup"><span data-stu-id="d91c8-312">1.1</span></span>|
|[<span data-ttu-id="d91c8-313">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d91c8-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d91c8-314">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d91c8-314">Compose or Read</span></span>|

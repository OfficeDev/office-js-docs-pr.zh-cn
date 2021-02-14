---
title: 在 Outlook 外接程序中启用委派访问方案
description: 简要介绍委派访问权限并讨论如何配置外接程序支持。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 598f931dbf3a4be8adf029838084ec0767bf6518
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234238"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="723b8-103">在 Outlook 外接程序中启用委派访问方案</span><span class="sxs-lookup"><span data-stu-id="723b8-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="723b8-104">邮箱所有者可以使用委派访问功能 [允许其他人管理其邮件和日历](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)。</span><span class="sxs-lookup"><span data-stu-id="723b8-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="723b8-105">本文指定 Office JavaScript API 支持哪些委派权限，并介绍如何在 Outlook 外接程序中启用委派访问方案。</span><span class="sxs-lookup"><span data-stu-id="723b8-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="723b8-106">代理访问当前在 Android 版 Outlook 和 iOS 中不可用。</span><span class="sxs-lookup"><span data-stu-id="723b8-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="723b8-107">此外，此功能当前不适用于 Outlook 网页 [中的组共享](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) 邮箱。</span><span class="sxs-lookup"><span data-stu-id="723b8-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="723b8-108">将来可能会提供此功能。</span><span class="sxs-lookup"><span data-stu-id="723b8-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="723b8-109">要求集 1.8 中引入了对此功能的支持。</span><span class="sxs-lookup"><span data-stu-id="723b8-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="723b8-110">请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="723b8-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="723b8-111">委派访问权限支持的权限</span><span class="sxs-lookup"><span data-stu-id="723b8-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="723b8-112">下表介绍了 Office JavaScript API 支持的委派权限。</span><span class="sxs-lookup"><span data-stu-id="723b8-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="723b8-113">权限</span><span class="sxs-lookup"><span data-stu-id="723b8-113">Permission</span></span>|<span data-ttu-id="723b8-114">值</span><span class="sxs-lookup"><span data-stu-id="723b8-114">Value</span></span>|<span data-ttu-id="723b8-115">说明</span><span class="sxs-lookup"><span data-stu-id="723b8-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="723b8-116">读取</span><span class="sxs-lookup"><span data-stu-id="723b8-116">Read</span></span>|<span data-ttu-id="723b8-117">1 (000001) </span><span class="sxs-lookup"><span data-stu-id="723b8-117">1 (000001)</span></span>|<span data-ttu-id="723b8-118">可读取项目。</span><span class="sxs-lookup"><span data-stu-id="723b8-118">Can read items.</span></span>|
|<span data-ttu-id="723b8-119">写入</span><span class="sxs-lookup"><span data-stu-id="723b8-119">Write</span></span>|<span data-ttu-id="723b8-120">2 (000010) </span><span class="sxs-lookup"><span data-stu-id="723b8-120">2 (000010)</span></span>|<span data-ttu-id="723b8-121">可以创建项目。</span><span class="sxs-lookup"><span data-stu-id="723b8-121">Can create items.</span></span>|
|<span data-ttu-id="723b8-122">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="723b8-122">DeleteOwn</span></span>|<span data-ttu-id="723b8-123">4 (000100) </span><span class="sxs-lookup"><span data-stu-id="723b8-123">4 (000100)</span></span>|<span data-ttu-id="723b8-124">只能删除他们创建的项。</span><span class="sxs-lookup"><span data-stu-id="723b8-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="723b8-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="723b8-125">DeleteAll</span></span>|<span data-ttu-id="723b8-126">8 (001000) </span><span class="sxs-lookup"><span data-stu-id="723b8-126">8 (001000)</span></span>|<span data-ttu-id="723b8-127">可以删除任何项目。</span><span class="sxs-lookup"><span data-stu-id="723b8-127">Can delete any items.</span></span>|
|<span data-ttu-id="723b8-128">EditOwn</span><span class="sxs-lookup"><span data-stu-id="723b8-128">EditOwn</span></span>|<span data-ttu-id="723b8-129">16 (010000) </span><span class="sxs-lookup"><span data-stu-id="723b8-129">16 (010000)</span></span>|<span data-ttu-id="723b8-130">只能编辑他们创建的项。</span><span class="sxs-lookup"><span data-stu-id="723b8-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="723b8-131">EditAll</span><span class="sxs-lookup"><span data-stu-id="723b8-131">EditAll</span></span>|<span data-ttu-id="723b8-132">32 (1000000) </span><span class="sxs-lookup"><span data-stu-id="723b8-132">32 (100000)</span></span>|<span data-ttu-id="723b8-133">可以编辑任何项目。</span><span class="sxs-lookup"><span data-stu-id="723b8-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="723b8-134">目前，API 支持获取现有委派权限，但不设置委派权限。</span><span class="sxs-lookup"><span data-stu-id="723b8-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="723b8-135">[DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)对象使用位掩码实现，以指示代理的权限。</span><span class="sxs-lookup"><span data-stu-id="723b8-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="723b8-136">位掩码中的每个位置表示特定权限，如果设置为该位置，则代理 `1` 具有各自的权限。</span><span class="sxs-lookup"><span data-stu-id="723b8-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="723b8-137">例如，如果右侧第二位为 `1` ，则代理具有 **写入** 权限。</span><span class="sxs-lookup"><span data-stu-id="723b8-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="723b8-138">您可以在本文稍后的"执行代理操作"部分查看如何检查特定权限[](#perform-an-operation-as-delegate)的示例。</span><span class="sxs-lookup"><span data-stu-id="723b8-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="723b8-139">跨邮箱客户端同步</span><span class="sxs-lookup"><span data-stu-id="723b8-139">Sync across mailbox clients</span></span>

<span data-ttu-id="723b8-140">代理对所有者邮箱的更新通常会立即跨邮箱同步。</span><span class="sxs-lookup"><span data-stu-id="723b8-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="723b8-141">但是，如果使用 REST 或 Exchange Web Services (EWS) 操作来设置项目的扩展属性，则此类更改可能需要几个小时才能同步。我们建议你改为使用 [CustomProperties](/javascript/api/outlook/office.customproperties) 对象和相关 API 以避免此类延迟。</span><span class="sxs-lookup"><span data-stu-id="723b8-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="723b8-142">若要了解更多信息，请参阅"[](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties)在 Outlook 外接程序中获取和设置元数据"一文的自定义属性部分。</span><span class="sxs-lookup"><span data-stu-id="723b8-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="723b8-143">在委托方案中，不能将 EWS 与当前由 office.js API 提供的令牌一同使用。</span><span class="sxs-lookup"><span data-stu-id="723b8-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="723b8-144">配置清单</span><span class="sxs-lookup"><span data-stu-id="723b8-144">Configure the manifest</span></span>

<span data-ttu-id="723b8-145">若要在外接程序中启用委派访问方案，必须在父元素下的清单中将 [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` 元素设置为 `DesktopFormFactor` 。</span><span class="sxs-lookup"><span data-stu-id="723b8-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="723b8-146">目前，不支持其他外形因素。</span><span class="sxs-lookup"><span data-stu-id="723b8-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="723b8-147">若要支持从代理进行 REST 调用，请设置清单中的 [Permissions](../reference/manifest/permissions.md) 节点 `ReadWriteMailbox` 。</span><span class="sxs-lookup"><span data-stu-id="723b8-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="723b8-148">下面的示例演示在 `SupportsSharedFolders` 清单的 `true` 一节中设置为的元素。</span><span class="sxs-lookup"><span data-stu-id="723b8-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="723b8-149">以委派方式执行操作</span><span class="sxs-lookup"><span data-stu-id="723b8-149">Perform an operation as delegate</span></span>

<span data-ttu-id="723b8-150">可以通过调用 [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法在撰写或阅读模式下获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="723b8-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="723b8-151">这将返回 [一个 SharedProperties](/javascript/api/outlook/office.sharedproperties) 对象，该对象当前提供代理的权限、所有者的电子邮件地址、REST API 的基本 URL 和目标邮箱。</span><span class="sxs-lookup"><span data-stu-id="723b8-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="723b8-152">以下示例演示如何获取邮件或约会的共享属性、检查代理是否具有写入权限以及进行 REST调用。</span><span class="sxs-lookup"><span data-stu-id="723b8-152">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> <span data-ttu-id="723b8-153">作为代理，可以使用 REST 获取附加到 Outlook 项目或组帖子的 [Outlook 邮件的内容](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)。</span><span class="sxs-lookup"><span data-stu-id="723b8-153">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="723b8-154">处理对共享和非共享项的调用 REST</span><span class="sxs-lookup"><span data-stu-id="723b8-154">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="723b8-155">如果要对项调用 REST 操作，无论该项目是否共享，都可以使用 API 来确定 `getSharedPropertiesAsync` 该项目是否共享。</span><span class="sxs-lookup"><span data-stu-id="723b8-155">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="723b8-156">之后，可以使用适当的对象构造该操作的 REST URL。</span><span class="sxs-lookup"><span data-stu-id="723b8-156">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a><span data-ttu-id="723b8-157">限制</span><span class="sxs-lookup"><span data-stu-id="723b8-157">Limitations</span></span>

<span data-ttu-id="723b8-158">根据加载项的方案，在处理委派情况时，需要考虑一些限制。</span><span class="sxs-lookup"><span data-stu-id="723b8-158">Depending on your add-in's scenarios, there are a couple of limitations for you to consider when handling delegate situations.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="723b8-159">REST 和 EWS</span><span class="sxs-lookup"><span data-stu-id="723b8-159">REST and EWS</span></span>

<span data-ttu-id="723b8-160">您的外接程序可以使用 REST，但不能使用 EWS，并且必须将外接程序的权限设置为启用对所有者邮箱 `ReadWriteMailbox` 的 REST 访问。</span><span class="sxs-lookup"><span data-stu-id="723b8-160">Your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="723b8-161">邮件撰写模式</span><span class="sxs-lookup"><span data-stu-id="723b8-161">Message Compose mode</span></span>

<span data-ttu-id="723b8-162">在邮件撰写模式下，除非满足以下条件，否则 Outlook 网页或 Windows 不支持[getSharedPropertiesAsync。](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-)</span><span class="sxs-lookup"><span data-stu-id="723b8-162">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) is not supported in Outlook on the web or Windows unless the following conditions are met.</span></span>

1. <span data-ttu-id="723b8-163">所有者至少与代理共享一个邮箱文件夹。</span><span class="sxs-lookup"><span data-stu-id="723b8-163">The owner shares at least one mailbox folder with the delegate.</span></span>
1. <span data-ttu-id="723b8-164">代理在共享文件夹中草稿邮件。</span><span class="sxs-lookup"><span data-stu-id="723b8-164">The delegate drafts a message in the shared folder.</span></span>

    <span data-ttu-id="723b8-165">示例：</span><span class="sxs-lookup"><span data-stu-id="723b8-165">Examples:</span></span>

    - <span data-ttu-id="723b8-166">代理答复或转发共享文件夹中的电子邮件。</span><span class="sxs-lookup"><span data-stu-id="723b8-166">The delegate replies to or forwards an email in the shared folder.</span></span>
    - <span data-ttu-id="723b8-167">然后，代理保存草稿邮件，然后将它从其自己的 **"草稿"** 文件夹移动到共享文件夹。</span><span class="sxs-lookup"><span data-stu-id="723b8-167">The delegate saves a draft message then moves it from their own **Drafts** folder to the shared folder.</span></span> <span data-ttu-id="723b8-168">代理从共享文件夹打开草稿，然后继续撰写。</span><span class="sxs-lookup"><span data-stu-id="723b8-168">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="723b8-169">邮件发送后，通常会在代理的"已发送项目"**文件夹中找到。**</span><span class="sxs-lookup"><span data-stu-id="723b8-169">After the message has been sent, it's usually found in the delegate's **Sent Items** folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="723b8-170">另请参阅</span><span class="sxs-lookup"><span data-stu-id="723b8-170">See also</span></span>

- [<span data-ttu-id="723b8-171">允许其他人管理邮件和日历</span><span class="sxs-lookup"><span data-stu-id="723b8-171">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="723b8-172">Microsoft 365 中的日历共享</span><span class="sxs-lookup"><span data-stu-id="723b8-172">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="723b8-173">如何对清单元素排序</span><span class="sxs-lookup"><span data-stu-id="723b8-173">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="723b8-174">[屏蔽 (计算) ](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="723b8-174">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="723b8-175">JavaScript 位运算符</span><span class="sxs-lookup"><span data-stu-id="723b8-175">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)
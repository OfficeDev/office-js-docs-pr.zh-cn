---
title: 在 Outlook 加载项中启用代理访问方案
description: 简要介绍了代理访问权限，并讨论了如何配置加载项支持。
ms.date: 09/30/2020
localization_priority: Normal
ms.openlocfilehash: 68e9c8003f8d223a591283fd1a73f0a38bd3c8a4
ms.sourcegitcommit: 6c3a04acde57832feeaaa599148f93af7e3e36ea
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/02/2020
ms.locfileid: "48336417"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="f6029-103">在 Outlook 加载项中启用代理访问方案</span><span class="sxs-lookup"><span data-stu-id="f6029-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="f6029-104">邮箱所有者可以使用代理访问功能，以 [允许其他人管理其邮件和日历](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)。</span><span class="sxs-lookup"><span data-stu-id="f6029-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="f6029-105">本文指定 Office JavaScript API 支持的代理权限，并介绍如何在 Outlook 外接程序中启用代理访问方案。</span><span class="sxs-lookup"><span data-stu-id="f6029-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f6029-106">目前在 Android 和 iOS 上的 Outlook 中不提供委派访问权限。</span><span class="sxs-lookup"><span data-stu-id="f6029-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="f6029-107">此外，此功能当前不适用于 web 上的 Outlook 中的 [组共享邮箱](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) 。</span><span class="sxs-lookup"><span data-stu-id="f6029-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="f6029-108">将来可提供此功能。</span><span class="sxs-lookup"><span data-stu-id="f6029-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="f6029-109">对此功能的支持是在要求集1.8 中引入的。</span><span class="sxs-lookup"><span data-stu-id="f6029-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="f6029-110">请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="f6029-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="f6029-111">代理访问支持的权限</span><span class="sxs-lookup"><span data-stu-id="f6029-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="f6029-112">下表介绍了 Office JavaScript API 支持的代理权限。</span><span class="sxs-lookup"><span data-stu-id="f6029-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="f6029-113">权限</span><span class="sxs-lookup"><span data-stu-id="f6029-113">Permission</span></span>|<span data-ttu-id="f6029-114">值</span><span class="sxs-lookup"><span data-stu-id="f6029-114">Value</span></span>|<span data-ttu-id="f6029-115">说明</span><span class="sxs-lookup"><span data-stu-id="f6029-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="f6029-116">阅读</span><span class="sxs-lookup"><span data-stu-id="f6029-116">Read</span></span>|<span data-ttu-id="f6029-117">1 (000001) </span><span class="sxs-lookup"><span data-stu-id="f6029-117">1 (000001)</span></span>|<span data-ttu-id="f6029-118">可以读取项目。</span><span class="sxs-lookup"><span data-stu-id="f6029-118">Can read items.</span></span>|
|<span data-ttu-id="f6029-119">写入</span><span class="sxs-lookup"><span data-stu-id="f6029-119">Write</span></span>|<span data-ttu-id="f6029-120">2 (000010) </span><span class="sxs-lookup"><span data-stu-id="f6029-120">2 (000010)</span></span>|<span data-ttu-id="f6029-121">可以创建项目。</span><span class="sxs-lookup"><span data-stu-id="f6029-121">Can create items.</span></span>|
|<span data-ttu-id="f6029-122">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="f6029-122">DeleteOwn</span></span>|<span data-ttu-id="f6029-123">4 (000100) </span><span class="sxs-lookup"><span data-stu-id="f6029-123">4 (000100)</span></span>|<span data-ttu-id="f6029-124">只能删除其创建的项目。</span><span class="sxs-lookup"><span data-stu-id="f6029-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="f6029-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="f6029-125">DeleteAll</span></span>|<span data-ttu-id="f6029-126">8 (001000) </span><span class="sxs-lookup"><span data-stu-id="f6029-126">8 (001000)</span></span>|<span data-ttu-id="f6029-127">可以删除任何项目。</span><span class="sxs-lookup"><span data-stu-id="f6029-127">Can delete any items.</span></span>|
|<span data-ttu-id="f6029-128">EditOwn</span><span class="sxs-lookup"><span data-stu-id="f6029-128">EditOwn</span></span>|<span data-ttu-id="f6029-129">16 (010000) </span><span class="sxs-lookup"><span data-stu-id="f6029-129">16 (010000)</span></span>|<span data-ttu-id="f6029-130">只能编辑其创建的项目。</span><span class="sxs-lookup"><span data-stu-id="f6029-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="f6029-131">EditAll</span><span class="sxs-lookup"><span data-stu-id="f6029-131">EditAll</span></span>|<span data-ttu-id="f6029-132">32 (100000) </span><span class="sxs-lookup"><span data-stu-id="f6029-132">32 (100000)</span></span>|<span data-ttu-id="f6029-133">可以编辑任何项目。</span><span class="sxs-lookup"><span data-stu-id="f6029-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="f6029-134">目前，API 支持获取现有的代理权限，但不支持设置委派权限。</span><span class="sxs-lookup"><span data-stu-id="f6029-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="f6029-135">使用位掩码来实现 [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) 对象，以指示代理的权限。</span><span class="sxs-lookup"><span data-stu-id="f6029-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="f6029-136">位掩码中的每个位置都代表一个特定权限，如果将其设置为，则 `1` 代理具有相应的权限。</span><span class="sxs-lookup"><span data-stu-id="f6029-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="f6029-137">例如，如果右边的第二位是 `1` ，则代理具有 " **写入** " 权限。</span><span class="sxs-lookup"><span data-stu-id="f6029-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="f6029-138">您可以在本文后面的 "将 [操作作为代理执行操作](#perform-an-operation-as-delegate) " 一节中查看有关如何检查特定权限的示例。</span><span class="sxs-lookup"><span data-stu-id="f6029-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="f6029-139">在邮箱客户端之间同步</span><span class="sxs-lookup"><span data-stu-id="f6029-139">Sync across mailbox clients</span></span>

<span data-ttu-id="f6029-140">代理对所有者邮箱的更新通常会在邮箱之间立即同步。</span><span class="sxs-lookup"><span data-stu-id="f6029-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="f6029-141">但是，如果 REST 或 Exchange Web 服务 (EWS) 操作用于设置项的扩展属性，则此类更改可能需要几个小时才能同步。我们建议您改为使用 [CustomProperties](/javascript/api/outlook/office.customproperties) 对象和相关 api 以避免此类延迟。</span><span class="sxs-lookup"><span data-stu-id="f6029-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="f6029-142">若要了解详细信息，请参阅 "在 Outlook 外接程序中获取和设置元数据" 一文中的 " [自定义属性" 部分](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) 。</span><span class="sxs-lookup"><span data-stu-id="f6029-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f6029-143">在委托方案中，不能将 EWS 与 office.js API 当前提供的令牌结合使用。</span><span class="sxs-lookup"><span data-stu-id="f6029-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="f6029-144">配置清单</span><span class="sxs-lookup"><span data-stu-id="f6029-144">Configure the manifest</span></span>

<span data-ttu-id="f6029-145">若要在外接程序中启用代理访问方案，必须在[SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` 父元素下的清单中将 SupportsSharedFolders 元素设置为 `DesktopFormFactor` 。</span><span class="sxs-lookup"><span data-stu-id="f6029-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="f6029-146">目前，其他外观因素不受支持。</span><span class="sxs-lookup"><span data-stu-id="f6029-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="f6029-147">若要支持来自代理的 REST 调用，请将清单中的 " [权限](../reference/manifest/permissions.md) " 节点设置为 `ReadWriteMailbox` 。</span><span class="sxs-lookup"><span data-stu-id="f6029-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="f6029-148">下面的示例演示 `SupportsSharedFolders` `true` 在清单的部分中设置的元素。</span><span class="sxs-lookup"><span data-stu-id="f6029-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="f6029-149">将操作作为代理执行</span><span class="sxs-lookup"><span data-stu-id="f6029-149">Perform an operation as delegate</span></span>

<span data-ttu-id="f6029-150">可以通过调用 [getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法，在撰写或阅读模式下获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="f6029-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="f6029-151">这将返回一个 [SharedProperties](/javascript/api/outlook/office.sharedproperties) 对象，该对象当前提供代理的权限、所有者的电子邮件地址、REST API 的基 URL 和目标邮箱。</span><span class="sxs-lookup"><span data-stu-id="f6029-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="f6029-152">下面的示例展示了如何获取邮件或约会的共享属性、检查代理是否具有 **写入** 权限，以及如何发出 REST 调用。</span><span class="sxs-lookup"><span data-stu-id="f6029-152">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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
> <span data-ttu-id="f6029-153">作为代理，您可以使用 REST [获取附加到 outlook 项目或组文章的 outlook 邮件的内容](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)。</span><span class="sxs-lookup"><span data-stu-id="f6029-153">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="f6029-154">处理共享和非共享项上的呼叫 REST</span><span class="sxs-lookup"><span data-stu-id="f6029-154">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="f6029-155">如果要对某个项目调用 REST 操作（无论该项目是否共享），则可以使用 `getSharedPropertiesAsync` API 来确定该项目是否已共享。</span><span class="sxs-lookup"><span data-stu-id="f6029-155">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="f6029-156">之后，可以使用相应的对象构造该操作的 REST URL。</span><span class="sxs-lookup"><span data-stu-id="f6029-156">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

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

## <a name="limitations"></a><span data-ttu-id="f6029-157">限制</span><span class="sxs-lookup"><span data-stu-id="f6029-157">Limitations</span></span>

<span data-ttu-id="f6029-158">根据你的外接程序的方案，在处理委派情况时需要考虑以下几个限制。</span><span class="sxs-lookup"><span data-stu-id="f6029-158">Depending on your add-in's scenarios, there are a couple of limitations for you to consider when handling delegate situations.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="f6029-159">REST 和 EWS</span><span class="sxs-lookup"><span data-stu-id="f6029-159">REST and EWS</span></span>

<span data-ttu-id="f6029-160">您的外接程序可以使用 REST 而不是 EWS，并且必须将外接程序的权限设置为，以 `ReadWriteMailbox` 启用对所有者邮箱的 rest 访问。</span><span class="sxs-lookup"><span data-stu-id="f6029-160">Your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="f6029-161">邮件撰写模式</span><span class="sxs-lookup"><span data-stu-id="f6029-161">Message Compose mode</span></span>

<span data-ttu-id="f6029-162">在邮件撰写模式下，在 web 或 Windows 上的 Outlook 中不支持 [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) ，除非满足以下条件。</span><span class="sxs-lookup"><span data-stu-id="f6029-162">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) is not supported in Outlook on the web or Windows unless the following conditions are met.</span></span>

1. <span data-ttu-id="f6029-163">所有者至少与委派共享一个邮箱文件夹。</span><span class="sxs-lookup"><span data-stu-id="f6029-163">The owner shares at least one mailbox folder with the delegate.</span></span>
1. <span data-ttu-id="f6029-164">代理将邮件草稿共享到共享文件夹中。</span><span class="sxs-lookup"><span data-stu-id="f6029-164">The delegate drafts a message in the shared folder.</span></span>

    <span data-ttu-id="f6029-165">示例：</span><span class="sxs-lookup"><span data-stu-id="f6029-165">Examples:</span></span>

    - <span data-ttu-id="f6029-166">代理答复或转发共享文件夹中的电子邮件。</span><span class="sxs-lookup"><span data-stu-id="f6029-166">The delegate replies to or forwards an email in the shared folder.</span></span>
    - <span data-ttu-id="f6029-167">代理保存草稿邮件，然后将其从其自己的 " **草稿** " 文件夹移动到共享文件夹。</span><span class="sxs-lookup"><span data-stu-id="f6029-167">The delegate saves a draft message then moves it from their own **Drafts** folder to the shared folder.</span></span> <span data-ttu-id="f6029-168">委派从共享文件夹打开草稿，然后继续撰写。</span><span class="sxs-lookup"><span data-stu-id="f6029-168">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="f6029-169">邮件发送后，通常会在代理的 " **已发送邮件** " 文件夹中找到。</span><span class="sxs-lookup"><span data-stu-id="f6029-169">After the message has been sent, it's usually found in the delegate's **Sent Items** folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="f6029-170">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f6029-170">See also</span></span>

- [<span data-ttu-id="f6029-171">允许其他人管理您的邮件和日历</span><span class="sxs-lookup"><span data-stu-id="f6029-171">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="f6029-172">Office 365 中的日历共享</span><span class="sxs-lookup"><span data-stu-id="f6029-172">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="f6029-173">如何对清单元素进行排序</span><span class="sxs-lookup"><span data-stu-id="f6029-173">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="f6029-174">[掩码 (计算) ](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="f6029-174">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="f6029-175">JavaScript 按位运算符</span><span class="sxs-lookup"><span data-stu-id="f6029-175">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)
---
title: 添加对 Outlook 外接程序的移动支持
description: 了解如何添加对 Outlook Mobile 的支持，包括如何更新外接程序清单并根据需要更改移动方案的代码。
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: c84b4aeb04cd2c8b3c2f0a7afa9fd1631c22afc5
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607540"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>添加对适用于 Outlook Mobile 的外接程序命令的支持

使用 Outlook Mobile 中的外接程序命令，用户可以访问与在 Outlook 网页版、Windows 和 Mac 中已有一些[限制](#code-considerations)) 相同的功能 (。 添加对 Outlook Mobile 的支持需要更新外接程序清单，并且可能会更改移动方案的代码。

## <a name="updating-the-manifest"></a>更新清单

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](/javascript/api/manifest/versionoverrides) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](/javascript/api/manifest/mobileformfactor).

This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.

以下示例显示元素中的 `MobileFormFactor` 单个任务窗格按钮。

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

这与 [DesktopFormFactor](/javascript/api/manifest/desktopformfactor) 元素中出现的元素非常相似，但具有一些明显的区别。

- 不使用 [OfficeTab](/javascript/api/manifest/officetab) 元素。
- The [ExtensionPoint](/javascript/api/manifest/extensionpoint) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](/javascript/api/manifest/control) element. If the add-in adds more than one button, the child element should be a [Group](/javascript/api/manifest/group) element that contains multiple `Control` elements.
- 没有与 `Control` 元素等效的 `Menu` 类型。
- 不使用 [Supertip](/javascript/api/manifest/supertip) 元素。
- The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.

## <a name="code-considerations"></a>代码注意事项

设计适用于移动电话的外接程序引入了一些额外注意事项。

### <a name="use-rest-instead-of-exchange-web-services"></a>使用 REST 代替 Exchange Web 服务

The [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.

邮箱要求集 1.5 引入了 [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 的新版本，可请求与 REST API 兼容的访问令牌，以及可用于查找用户 REST API 终结点的新 [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) 属性。

### <a name="pinch-zoom"></a>收缩缩放

By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.

### <a name="close-task-panes"></a>关闭任务窗格

In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) method to close the task pane when your scenario is complete.

### <a name="compose-mode-and-appointments"></a>撰写模式和约会

目前，Outlook Mobile 中的加载项仅支持在读取消息时激活。 在撰写邮件时或查看或撰写约会时，不会激活外接程序。 但是，有两个例外：

1. 可在约会组织者模式下激活联机会议提供商集成的加载项。 有关此异常 (包括可用 API) 的详细信息，请参阅 [为联机会议提供商创建 Outlook 移动外接程序](online-meeting.md#available-apis)。
1. 可在约会与会者模式下激活将约会笔记和其他详细信息记录到客户关系管理 (CRM) 或记事服务的加载项。 有关此异常 (包括可用 API) 的详细信息，请参阅 [Outlook 移动外接程序中的外部应用程序的日志约会说明](mobile-log-appointments.md#available-apis)。

### <a name="unsupported-apis"></a>不支持的 API

Outlook Mobile 不支持要求集 1.6 或更高版本中引入的 API。 也不支持来自早期要求集的以下 API。

- [Office.context.officeTheme](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context#officetheme-officetheme)
- [Office.context.mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
- [Office.context.mailbox.convertToEwsId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayMessageForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
- [Office.context.mailbox.item.displayReplyAllForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.displayReplyForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntitiesByType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

## <a name="see-also"></a>另请参阅

[Exchange 服务器和 Outlook 客户端支持的要求集](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
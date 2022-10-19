---
title: Outlook 外接程序清单
description: 获取 Outlook 外接程序可用的两种清单的概述。
ms.date: 10/18/2022
ms.localizationpriority: high
ms.openlocfilehash: a22b5180fee6b4f9f0663eff54b57510016202a2
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607553"
---
# <a name="outlook-add-in-manifests"></a>Outlook 外接程序清单

Outlook 外接程序由两个组件组成：外接程序清单和 Office 外接程序的 JavaScript 库 (office.js) 支持的 Web 应用。 清单介绍了外接程序如何跨 Outlook 客户端集成。

清单有两种可能的格式：XML 和 JSON。 可以了解 [Office 外接程序的 Teams 清单中的 JSON 清单 (预览) ](../develop/json-manifest-overview.md)。 本文介绍 XML 清单。

下面是 XML 清单的示例。

 > [!NOTE]
 > All URL values in the following sample begin with "https://appdemo.contoso.com". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.

```XML
<?xml version="1.0" encoding="UTF-8" ?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft Outlook Dev Center</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add-in Command Demo" />
  <Description DefaultValue="Adds command buttons to the ribbon in Outlook"/>
  <IconUrl DefaultValue="https://appdemo.contoso.com/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="https://appdemo.contoso.com/images/blue-128.png" />
  <SupportUrl DefaultValue="https://appdemo.contoso.com"/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <!-- These elements support older clients that don't support add-in commands -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- NOTE: Just reusing the read task pane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="functionFile" />

          <!-- Message read form -->
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadDemoGroup">
                <Label resid="groupLabel" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="msgReadFunctionButton">
                  <Label resid="funcReadButtonLabel" />
                  <Supertip>
                    <Title resid="funcReadSuperTipTitle" />
                    <Description resid="funcReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getSubject</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="msgReadMenuButton">
                  <Label resid="menuReadButtonLabel" />
                  <Supertip>
                    <Title resid="menuReadSuperTipTitle" />
                    <Description resid="menuReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="msgReadMenuItem1">
                      <Label resid="menuItem1ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem1ReadLabel" />
                        <Description resid="menuItem1ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemClass</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem2">
                      <Label resid="menuItem2ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem2ReadLabel" />
                        <Description resid="menuItem2ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getDateTimeCreated</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem3">
                      <Label resid="menuItem3ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem3ReadLabel" />
                        <Description resid="menuItem3ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemID</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="paneReadButtonLabel" />
                  <Supertip>
                    <Title resid="paneReadSuperTipTitle" />
                    <Description resid="paneReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="readTaskPaneUrl" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <!-- Blue icon -->
        <bt:Image id="blue-icon-16" DefaultValue="https://appdemo.contoso.com/images/blue-16.png" />
        <bt:Image id="blue-icon-32" DefaultValue="https://appdemo.contoso.com/images/blue-32.png" />
        <bt:Image id="blue-icon-80" DefaultValue="https://appdemo.contoso.com/images/blue-80.png" />
        <!-- Red icon -->
        <bt:Image id="red-icon-16" DefaultValue="https://appdemo.contoso.com/images/red-16.png" />
        <bt:Image id="red-icon-32" DefaultValue="https://appdemo.contoso.com/images/red-32.png" />
        <bt:Image id="red-icon-80" DefaultValue="https://appdemo.contoso.com/images/red-80.png" />
        <!-- Green icon -->
        <bt:Image id="green-icon-16" DefaultValue="https://appdemo.contoso.com/images/green-16.png" />
        <bt:Image id="green-icon-32" DefaultValue="https://appdemo.contoso.com/images/green-32.png" />
        <bt:Image id="green-icon-80" DefaultValue="https://appdemo.contoso.com/images/green-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="https://appdemo.contoso.com/FunctionFile/Functions.html" />
        <bt:Url id="readTaskPaneUrl" DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="Add-in Demo" />
        <bt:String id="funcReadButtonLabel" DefaultValue="Get subject" />
        <bt:String id="menuReadButtonLabel" DefaultValue="Get property" />
        <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties" />

        <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment" />
        <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get" />
        <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties" />

        <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class" />
        <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created" />
        <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button." />
        <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button." />
        <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane." />

        <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
        <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
        <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

## <a name="schema-versions"></a>架构版本

Not all Outlook clients support the latest features, and some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.

清单中的 **\<VersionOverrides\>** 元素就是此类情况的示例。 **\<VersionOverrides\>** 中定义的所有元素将替代清单另一部分中的同一元素。 这意味着，只要有可能，Outlook 就会使用 **\<VersionOverrides\>** 部分中的内容设置加载项。 但是，如果 Outlook 版本不支持 **\<VersionOverrides\>** 的某个版本，Outlook 则会将其忽略，具体取决于清单其余部分中的信息。 

此方法意味着开发人员无需创建多个单独的清单，而是将定义的所有内容保留在一个文件中。

架构的当前版本为：


|版本|说明|
|:-----|:-----|
|v1.0|Supports version 1.0 of the Office JavaScript API. For Outlook add-ins, this supports read form. |
|v1.1|支持 Office JavaScript API 版本 1.1 和 **\<VersionOverrides\>**。 对于 Outlook 外接程序，它将添加对撰写窗体的支持。|
|**\<VersionOverrides\>** 1.0|支持 Office JavaScript API 的更高版本。 这支持外接程序命令。|
|**\<VersionOverrides\>** 1.1|Supports later versions of the Office JavaScript API. This supports add-in commands and adds support for newer features, such as [pinnable task panes](pinnable-taskpane.md) and mobile add-ins.|

本文将介绍 v1.1 清单的要求。 即使加载项清单使用 **\<VersionOverrides\>** 元素，仍需将 v1.1 清单元素包括在内，以允许加载项使用不支持 **\<VersionOverrides\>** 的旧版客户端。

> [!NOTE]
> Outlook uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. You can download the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to help create your manifest with elements in the required order.

## <a name="root-element"></a>根元素

Outlook 加载项清单的根元素是 **\<OfficeApp\>**。 此元素还声明默认命名空间、架构版本和外接程序类型。 将清单中的所有其他元素置于其开始标记和结束标记中。 下面是根元素的一个示例。


```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest -->

</OfficeApp>
```

## <a name="version"></a>版本

This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.

If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will continue to see old functionality in the meantime.

## <a name="versionoverrides"></a>VersionOverrides

**\<VersionOverrides\>** 元素是 [加载项命令](add-in-commands-for-outlook.md) 的信息的位置。

此元素也是外接程序为[移动外接程序](add-mobile-support.md)定义支持所使用的元素。

有关此元素的讨论，请参阅[在清单中创建 Excel、PowerPoint 和 Word 加载项命令](../develop/create-addin-commands.md)。

## <a name="localization"></a>本地化

外接程序的某些方面需要进行本地化以适用于不同的区域设置，例如名称、介绍以及所加载的 URL。 可通过指定默认值，然后在 **\<VersionOverrides\>** 元素的 **\<Resources\>** 元素中指定区域设置替代来轻松地实现这些元素的本地化。 下面介绍了如何替代图像、URL 和字符串。


```XML
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
      <!-- add information for other locales -->
    </bt:Image>
  </bt:Images>

  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
      <!-- add information for other locales -->
    </bt:Url>
  </bt:Urls>

  <bt:ShortStrings> 
    <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
      <bt:Override Locale="ar-sa" Value="<add localized value here>" />
      <!-- add information for other locales -->
    </bt:String>
  </bt:ShortStrings>
</Resources>
```

架构引用包含可本地化的元素的完整信息。

## <a name="hosts"></a>Hosts

Outlook 加载项指定如下所示的 **\<Hosts\>** 元素：

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

这与 **\<VersionOverrides\>** 元素中的 **\<Hosts\>** 元素有所不同，后者将在 [在清单中为 Excel、PowerPoint 和 Word 创建加载项命令](../develop/create-addin-commands.md) 中进行讨论。

## <a name="requirements"></a>要求

**\<Requirements\>** 元素指定可用于加载项的 API 集。 对于 Outlook 外接程序，要求集必须是邮箱和 1.1 或以上的值。 请参阅最新要求集版本的 API 引用。 有关要求集的详细信息，请参阅 [Outlook 外接程序 API](apis.md)。

**\<Requirements\>** 元素也可能出现在 **\<VersionOverrides\>** 元素中，因此加载项可以在加载到支持 **\<VersionOverrides\>** 的客户端中时指定不同的要求。

下面的示例使用 **\<Sets\>** 元素的 **DefaultMinVersion** 属性来要求 office.js 版本 1.1 或更高版本，并使用 **\<Set\>** 元素的 **MinVersion** 属性来要求邮箱要求集版本 1.1。

```XML
<OfficeApp>
...
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
...
</OfficeApp>
```

## <a name="form-settings"></a>表单设置

旧版 Outlook 客户端使用 **\<FormSettings\>** 元素，这仅支持架构 1.1，而不支持 **\<VersionOverrides\>**。 使用此元素，开发人员可以定义外接程序在此类客户端中显示的方式。 包含两个部分 - **ItemRead** 和 **ItemEdit**。 **ItemRead** 用于指定当用户阅读邮件和约会时外接程序的显示方式。 **ItemEdit** 说明当用户在撰写回复、新邮件、新约会或用户作为组织者编辑约会时外接程序的显示方式。

这些设置与 **\<Rule\>** 元素中的激活规则直接相关。 如果外接程序指定其应出现在撰写模式下的邮件中，则必须指定一个 **ItemEdit** 窗体。

有关更多详细信息，请参阅 Schema reference for Office Add-ins manifests (v1.1)。

## <a name="app-domains"></a>应用域

在 **\<SourceLocation\>** 元素中指定的加载项起始页的域为该加载项的默认域。 在未使用 **\<AppDomains\>** 和 **\<AppDomain\>** 元素的情况下，如果加载项尝试导航到其他域，浏览器将在加载项窗格以外打开一个新窗口。 要允许加载项导航到加载项窗格中的另一个域，请在加载项清单中添加 **\<AppDomains\>** 元素，并在其自己的 **\<AppDomain\>** 子元素中包括其他每个域。

以下示例指定域  `https://www.contoso2.com` 作为外接程序可以在外接程序窗格内导航到的第二个域。

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

对于在弹出窗口与运行在富客户端中的外接程序之间启用 cookie 共享而言，应用程序域也是必须的。

下表描述了浏览器在加载项尝试导航至加载项默认域外部 URL 时的行为。

|Outlook 客户端|已定义的域<br>是否在 AppDomains 中？|浏览器行为|
|---|---|---|
|所有客户端|是|链接将在加载项任务窗格中打开。|
|- 在 Windows 上Outlook 2016 (批量许可的永久) <br>- Windows 上的 Outlook 2013 (永久) |否|链接将在 Internet Explorer 11 中打开。|
|其他客户端|否|链接将在用户的默认浏览器中打开。|

有关更多详细信息，请参阅[指定要在加载项窗口中打开的域](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window)。

## <a name="permissions"></a>权限

**\<Permissions\>** 元素包含加载项所需的权限。 通常情况下，你应指定外接程序所需的最低权限，这取决于你计划使用的具体方法。 例如，如果在撰写窗体中激活的邮件外接程序仅读取但不写入 [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 等项目属性，也不调用 [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 访问任何 Exchange Web 服务操作，则应指定 **ReadItem** 权限。 有关可用权限的详细信息，请参阅[了解 Outlook 外接程序的权限](understanding-outlook-add-in-permissions.md)。

**邮件外接程序的 4 层权限模型**

![邮件应用架构 v1.1 的 4 层权限模型。](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## <a name="activation-rules"></a>激活规则

激活规则在 **\<Rule\>** 元素中指定。 **\<Rule\>** 元素可显示为 1.1 清单中 **\<OfficeApp\>** 元素的子元素。

激活规则可用于根据当前所选项目的下列一个或多个条件激活外接程序。

> [!NOTE]
> 激活规则只适用于不支持 **\<VersionOverrides\>** 元素的客户端。

- 项目类型和/或邮件类别

- 存在特定类型的已知实体，例如地址或电话号码

- 正文、主题或发件人电子邮件地址中的正则表达式匹配

- 存在附件

有关激活规则的详细信息和示例，请参阅 [Outlook 外接程序的激活规则](activation-rules.md)。

## <a name="next-steps-add-in-commands"></a>后续步骤：外接程序命令

After defining a basic manifest, define add-in commands for your add-in. Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).

有关定义外接程序命令的示例外接程序，请参阅 [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo)。

## <a name="next-steps-add-mobile-support"></a>后续步骤：添加移动支持

Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and Mac. For more information, see [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).

## <a name="see-also"></a>另请参阅

- [Office 外接程序的本地化](../develop/localization.md)
- [Outlook 外接程序的隐私、权限和安全性](privacy-and-security.md)
- [Outlook 外接程序 API](apis.md)
- [Office 外接程序 XML 清单](../develop/add-in-manifests.md)
- [Office 外接程序清单的架构参考 (v1.1)](../develop/add-in-manifests.md)
- [设计 Office 加载项](../design/add-in-design.md)
- [了解 Outlook 外接程序权限](understanding-outlook-add-in-permissions.md)
- [使用正则表达式激活规则显示 Outlook 外接程序](use-regular-expressions-to-show-an-outlook-add-in.md)
- [将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)
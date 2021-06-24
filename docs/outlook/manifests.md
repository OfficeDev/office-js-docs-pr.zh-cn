---
title: Outlook 外接程序清单
description: 该清单介绍 Outlook 外接程序如何跨 Outlook 客户端进行集成；其中包括一个示例。
ms.date: 05/27/2020
localization_priority: Priority
ms.openlocfilehash: f113a5d8f92ee80ed635283e9e5544bd4b9ce7cd
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076768"
---
# <a name="outlook-add-in-manifests"></a><span data-ttu-id="7620a-103">Outlook 外接程序清单</span><span class="sxs-lookup"><span data-stu-id="7620a-103">Outlook add-in manifests</span></span>

<span data-ttu-id="7620a-p101">Outlook 外接程序包括两个组件：XML 外接程序清单和网页，它们由 Office 外接程序的 JavaScript 库 (office.js) 提供支持。该清单介绍了外接程序如何跨 Outlook 客户端进行集成。示例如下。</span><span class="sxs-lookup"><span data-stu-id="7620a-p101">An Outlook add-in consists of two components: the XML add-in manifest and a web page supported by the JavaScript library for Office Add-ins (office.js). The manifest describes how the add-in integrates across Outlook clients. The following is an example.</span></span>

 > [!NOTE]
 > <span data-ttu-id="7620a-p102">以下示例中的所有 URL 值均以“https://appdemo.contoso.com”开头。该值是一个占位符。在实际的有效清单中，这些值将包含有效的 https Web URL。</span><span class="sxs-lookup"><span data-stu-id="7620a-p102">All URL values in the following sample begin with "https://appdemo.contoso.com". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.</span></span>

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

## <a name="schema-versions"></a><span data-ttu-id="7620a-110">架构版本</span><span class="sxs-lookup"><span data-stu-id="7620a-110">Schema versions</span></span>

<span data-ttu-id="7620a-p103">并非所有 Outlook 客户端均支持最新功能，某些 Outlook 用户可能使用的是旧版本的 Outlook。通过架构版本，开发人员可以使用可用的最新功能生成向后兼容的外接程序，同时仍能在旧版本上正常工作。</span><span class="sxs-lookup"><span data-stu-id="7620a-p103">Not all Outlook clients support the latest features, and some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.</span></span>

<span data-ttu-id="7620a-p104">清单中的 **VersionOverrides** 元素是此类型的一个示例。**VersionOverrides** 中定义的所有元素将都重写清单另一部分中的同一元素。这意味着，只要有可能，Outlook 都将使用 **VersionOverrides** 部分中的内容设置加载项。但是，如果 Outlook 版本不支持 **VersionOverrides** 的某个版本，Outlook 会将其忽略，具体取决于清单其余部分中的信息。</span><span class="sxs-lookup"><span data-stu-id="7620a-p104">The **VersionOverrides** element in the manifest is an example of this. All elements defined inside **VersionOverrides** will override the same element in the other part of the manifest. This means that, whenever possible, Outlook will use what is in the **VersionOverrides** section to set up the add-in. However, if the version of Outlook doesn't support a certain version of **VersionOverrides**, Outlook will ignore it and depend on the information in the rest of the manifest.</span></span> 

<span data-ttu-id="7620a-117">此方法意味着开发人员无需创建多个单独的清单，而是将定义的所有内容保留在一个文件中。</span><span class="sxs-lookup"><span data-stu-id="7620a-117">This approach means that developers don't have to create multiple individual manifests, but rather keep everything defined in one file.</span></span>

<span data-ttu-id="7620a-118">架构的当前版本为：</span><span class="sxs-lookup"><span data-stu-id="7620a-118">The current versions of the schema are:</span></span>


|<span data-ttu-id="7620a-119">版本</span><span class="sxs-lookup"><span data-stu-id="7620a-119">Version</span></span>|<span data-ttu-id="7620a-120">说明</span><span class="sxs-lookup"><span data-stu-id="7620a-120">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="7620a-121">v1.0</span><span class="sxs-lookup"><span data-stu-id="7620a-121">v1.0</span></span>|<span data-ttu-id="7620a-p105">支持 Office JavaScript API 版本 1.0。对于 Outlook 外接程序，它支持阅读窗体。</span><span class="sxs-lookup"><span data-stu-id="7620a-p105">Supports version 1.0 of the Office JavaScript API. For Outlook add-ins, this supports read form.</span></span> |
|<span data-ttu-id="7620a-124">v1.1</span><span class="sxs-lookup"><span data-stu-id="7620a-124">v1.1</span></span>|<span data-ttu-id="7620a-p106">支持 Office JavaScript API 版本 1.1 和 **VersionOverrides**。对于 Outlook 外接程序，现已开始支持撰写窗体。</span><span class="sxs-lookup"><span data-stu-id="7620a-p106">Supports version 1.1 of the Office JavaScript API and **VersionOverrides**. For Outlook add-ins, this adds support for compose form.</span></span>|
|<span data-ttu-id="7620a-127">**VersionOverrides** 1.0</span><span class="sxs-lookup"><span data-stu-id="7620a-127">**VersionOverrides** 1.0</span></span>|<span data-ttu-id="7620a-p107">支持 Office JavaScript API 的更高版本。这支持外接程序命令。</span><span class="sxs-lookup"><span data-stu-id="7620a-p107">Supports later versions of the Office JavaScript API. This supports add-in commands.</span></span>|
|<span data-ttu-id="7620a-130">**VersionOverrides** 1.1</span><span class="sxs-lookup"><span data-stu-id="7620a-130">**VersionOverrides** 1.1</span></span>|<span data-ttu-id="7620a-p108">支持 Office JavaScript API 的更高版本。这支持外接程序命令并添加了对较新功能的支持，如[可固定的任务窗格](pinnable-taskpane.md)和移动外接程序。</span><span class="sxs-lookup"><span data-stu-id="7620a-p108">Supports later versions of the Office JavaScript API. This supports add-in commands and adds support for newer features, such as [pinnable task panes](pinnable-taskpane.md) and mobile add-ins.</span></span>|

<span data-ttu-id="7620a-p109">本文将介绍 v1.1 清单的要求。即使你的加载项清单使用 **VersionOverrides** 元素，仍需将 v1.1 清单元素包括在内，以允许加载项使用不支持 **VersionOverrides** 的旧版客户端。</span><span class="sxs-lookup"><span data-stu-id="7620a-p109">This article will cover the requirements for a v1.1 manifest. Even if your add-in manifest uses the **VersionOverrides** element, it is still important to include the v1.1 manifest elements to allow your add-in to work with older clients that do not support **VersionOverrides**.</span></span>

> [!NOTE]
> <span data-ttu-id="7620a-p110">Outlook 使用架构来验证清单。此架构要求清单中的元素按特定顺序显示。如果未按规定顺序添加元素，可能会在旁加载加载项时出现错误。可下载 [XML 架构定义 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)，帮助创建所含元素按规定顺序排列的清单。</span><span class="sxs-lookup"><span data-stu-id="7620a-p110">Outlook uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. You can download the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to help create your manifest with elements in the required order.</span></span>

## <a name="root-element"></a><span data-ttu-id="7620a-139">根元素</span><span class="sxs-lookup"><span data-stu-id="7620a-139">Root element</span></span>

<span data-ttu-id="7620a-p111">Outlook 外接程序清单的根元素是 **OfficeApp**。此元素还声明默认命名空间、架构版本和外接程序类型。请将清单中的其他所有元素都置于开始标记和结束标记内。根元素示例如下：</span><span class="sxs-lookup"><span data-stu-id="7620a-p111">The root element for the Outlook add-in manifest is **OfficeApp**. This element also declares the default namespace, schema version and the type of add-in. Place all other elements in the manifest within its open and close tags. The following is an example of the root element:</span></span>


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

## <a name="version"></a><span data-ttu-id="7620a-144">Version</span><span class="sxs-lookup"><span data-stu-id="7620a-144">Version</span></span>

<span data-ttu-id="7620a-p112">这是特定外接程序的版本。如果开发人员更新清单中的某些内容，版本也必须随之递增。因此，在安装新清单时，它将覆盖现有清单，并且用户将获得新功能。如果已将此外接程序提交至应用商店，则必须重新提交新清单并重新验证。然后，此外接程序的用户将在该清单被批准后几小时内自动获得新更新的清单。</span><span class="sxs-lookup"><span data-stu-id="7620a-p112">This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.</span></span>

<span data-ttu-id="7620a-p113">如果外接程序所请求的权限发生了更改，则系统将提示用户对外接程序进行升级和重新许可。如果管理员为整个组织安装该外接程序，则管理员需首先对其重新许可。同时，用户将继续看到旧功能。</span><span class="sxs-lookup"><span data-stu-id="7620a-p113">If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will continue to see old functionality in the meantime.</span></span>

## <a name="versionoverrides"></a><span data-ttu-id="7620a-153">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="7620a-153">VersionOverrides</span></span>

<span data-ttu-id="7620a-154">**VersionOverrides** 元素是 [外接程序命令](add-in-commands-for-outlook.md)信息的位置。</span><span class="sxs-lookup"><span data-stu-id="7620a-154">The **VersionOverrides** element is the location of information for [add-in commands](add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="7620a-155">此元素也是外接程序为[移动外接程序](add-mobile-support.md)定义支持所使用的元素。</span><span class="sxs-lookup"><span data-stu-id="7620a-155">This element is also where add-ins define support for [mobile add-ins](add-mobile-support.md).</span></span>

<span data-ttu-id="7620a-156">有关此元素的讨论，请参阅[在清单中创建 Excel、PowerPoint 和 Word 加载项命令](../develop/create-addin-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="7620a-156">For a discussion on this element, see [Create add-in commands in your manifest for Excel, PowerPoint, and Word](../develop/create-addin-commands.md).</span></span>

## <a name="localization"></a><span data-ttu-id="7620a-157">本地化</span><span class="sxs-lookup"><span data-stu-id="7620a-157">Localization</span></span>

<span data-ttu-id="7620a-p114">加载项的某些方面需要进行本地化以适用于不同的区域设置，例如名称、介绍以及所加载的 URL。可通过指定默认值并在 **VersionOverrides** 元素内的 **Resources** 元素中进行局部替换来轻松地实现这些元素的本地化。下面介绍了如何替换图像、URL 和字符串：</span><span class="sxs-lookup"><span data-stu-id="7620a-p114">Some aspects of the add-in need to be localized for different locales, such as the name, description and the URL that's loaded. These elements can easily be localized by specifying the default value and then locale overrides in the **Resources** element within the **VersionOverrides** element. The following shows how to override an image, a URL, and a string:</span></span>


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

<span data-ttu-id="7620a-161">架构引用包含可本地化的元素的完整信息。</span><span class="sxs-lookup"><span data-stu-id="7620a-161">The schema reference contains full information on which elements can be localized.</span></span>

## <a name="hosts"></a><span data-ttu-id="7620a-162">Hosts</span><span class="sxs-lookup"><span data-stu-id="7620a-162">Hosts</span></span>

<span data-ttu-id="7620a-163">Outlook 加载项指定如下所示的 **Hosts** 元素。</span><span class="sxs-lookup"><span data-stu-id="7620a-163">Outlook add-ins specify the **Hosts** element like the following.</span></span>

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

<span data-ttu-id="7620a-164">这与 **VersionOverrides** 元素中的 **Hosts** 元素有所不同，后者将在 [在清单中为 Excel、PowerPoint 和 Word 创建加载项命令](../develop/create-addin-commands.md)中进行讨论。</span><span class="sxs-lookup"><span data-stu-id="7620a-164">This is separate from the **Hosts** element inside the **VersionOverrides** element, which is discussed in [Create add-in commands in your manifest for Excel, PowerPoint, and Word](../develop/create-addin-commands.md).</span></span>

## <a name="requirements"></a><span data-ttu-id="7620a-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="7620a-165">Requirements</span></span>

<span data-ttu-id="7620a-p115">**Requirements** 元素指定外接程序可用的 API 集。对于 Outlook 外接程序，要求集必须是邮箱版本 1.1 或更高版本。请参阅最新要求集版本的 API 参考。若要详细了解要求集，请参阅 [Outlook 外接程序 API](apis.md)。</span><span class="sxs-lookup"><span data-stu-id="7620a-p115">The **Requirements** element specifies the set of APIs available to the add-in. For an Outlook add-in, the requirement set must be Mailbox and a value of 1.1 or above. Please refer to the API reference for the latest requirement set version. Refer to the [Outlook add-in APIs](apis.md) for more information on requirement sets.</span></span>

<span data-ttu-id="7620a-170">**Requirements** 元素也可能出现在 **VersionOverrides** 元素中，因此加载项可以在加载到支持 **VersionOverrides** 的客户端中时指定不同的要求。</span><span class="sxs-lookup"><span data-stu-id="7620a-170">The **Requirements** element can also appear in the **VersionOverrides** element, allowing the add-in to specify a different requirement when loaded in clients that support **VersionOverrides**.</span></span>

<span data-ttu-id="7620a-171">下面的示例使用 **Sets** 元素的 **DefaultMinVersion** 属性要求 office.js 版本 1.1 或更高版本，使用 **Set** 元素的 **MinVersion** 属性要求邮箱要求集版本 1.1。</span><span class="sxs-lookup"><span data-stu-id="7620a-171">The following example uses the **DefaultMinVersion** attribute of the **Sets** element to require office.js version 1.1 or higher, and the **MinVersion** attribute of the **Set** element to require the Mailbox requirement set version 1.1.</span></span>

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

## <a name="form-settings"></a><span data-ttu-id="7620a-172">表单设置</span><span class="sxs-lookup"><span data-stu-id="7620a-172">Form settings</span></span>

<span data-ttu-id="7620a-p116">旧版 Outlook 客户端使用的 **FormSettings** 元素仅支持架构 1.1，而不支持 **VersionOverrides**。使用此元素，开发人员可以定义加载项在此类客户端中显示的方式。包含两个部分：**ItemRead** 和 **ItemEdit**。**ItemRead** 用于指定当用户阅读邮件和约会时加载项的显示方式。**ItemEdit** 说明当用户在撰写回复、新邮件、新约会或用户作为组织者编辑约会时加载项的显示方式。</span><span class="sxs-lookup"><span data-stu-id="7620a-p116">The **FormSettings** element is used by older Outlook clients, which only support schema 1.1 and not **VersionOverrides**. Using this element, developers define how the add-in will appear in such clients. There are two parts - **ItemRead** and **ItemEdit**. **ItemRead** is used to specify how the add-in appears when the user reads messages and appointments. **ItemEdit** describes how the add-in appears while the user is composing a reply, new message, new appointment or editing an appointment where they are the organizer.</span></span>

<span data-ttu-id="7620a-p117">这些设置与 **Rule** 元素中的激活规则直接相关。例如，如果加载项指定其应出现在撰写模式下的邮件中，则必须指定一个 **ItemEdit** 窗体。</span><span class="sxs-lookup"><span data-stu-id="7620a-p117">These settings are directly related to the activation rules in the **Rule** element. For example, if an add-in specifies that it should appear on a message in compose mode, an **ItemEdit** form must be specified.</span></span>

<span data-ttu-id="7620a-180">有关更多详细信息，请参阅 Schema reference for Office Add-ins manifests (v1.1)。</span><span class="sxs-lookup"><span data-stu-id="7620a-180">For more details, please refer to the [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>

## <a name="app-domains"></a><span data-ttu-id="7620a-181">应用域</span><span class="sxs-lookup"><span data-stu-id="7620a-181">App domains</span></span>

<span data-ttu-id="7620a-p118">在 **SourceLocation** 元素中指定的加载项起始页的域为该上下文的默认域。在未使用 **AppDomains** 和 **AppDomain** 元素的情况下，如果加载项尝试导航到其他域，浏览器将在加载项窗格以外打开一个新窗口。要允许加载项导航到加载项窗格中的另一个域，请在加载项清单中添加 **AppDomains** 元素，并在其自己的 **AppDomain** 子元素中添加其他每个域。</span><span class="sxs-lookup"><span data-stu-id="7620a-p118">The domain of the add-in start page that you specify in the **SourceLocation** element is the default domain for the add-in. Without using the **AppDomains** and **AppDomain** elements, if your add-in attempts to navigate to another domain, the browser will open a new window outside of the add-in pane. In order to allow the add-in to navigate to another domain within the add-in pane, add an **AppDomains** element and include each additional domain in its own **AppDomain** sub-element in the add-in manifest.</span></span>

<span data-ttu-id="7620a-185">以下示例指定域  `https://www.contoso2.com` 作为外接程序可以在外接程序窗格内导航到的第二个域：</span><span class="sxs-lookup"><span data-stu-id="7620a-185">The following example specifies a domain  `https://www.contoso2.com` as a second domain that the add-in can navigate to within the add-in pane:</span></span>

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

<span data-ttu-id="7620a-186">对于在弹出窗口与运行在富客户端中的外接程序之间启用 cookie 共享而言，应用程序域也是必须的。</span><span class="sxs-lookup"><span data-stu-id="7620a-186">App domains are also necessary to enable cookie sharing between the pop-out window and the add-in running in the rich client.</span></span>

<span data-ttu-id="7620a-187">下表描述了浏览器在加载项尝试导航至加载项默认域外部 URL 时的行为。</span><span class="sxs-lookup"><span data-stu-id="7620a-187">The following table describes browser behavior when your add-in attempts to navigate to a URL outside of the add-in's default domain.</span></span>

|<span data-ttu-id="7620a-188">Outlook 客户端</span><span class="sxs-lookup"><span data-stu-id="7620a-188">Outlook client</span></span>|<span data-ttu-id="7620a-189">已定义的域</span><span class="sxs-lookup"><span data-stu-id="7620a-189">Domain defined</span></span><br><span data-ttu-id="7620a-190">是否在 AppDomains 中？</span><span class="sxs-lookup"><span data-stu-id="7620a-190">in AppDomains?</span></span>|<span data-ttu-id="7620a-191">浏览器行为</span><span class="sxs-lookup"><span data-stu-id="7620a-191">Browser behavior</span></span>|
|---|---|---|
|<span data-ttu-id="7620a-192">所有客户端</span><span class="sxs-lookup"><span data-stu-id="7620a-192">All clients</span></span>|<span data-ttu-id="7620a-193">是</span><span class="sxs-lookup"><span data-stu-id="7620a-193">Yes</span></span>|<span data-ttu-id="7620a-194">链接将在加载项任务窗格中打开。</span><span class="sxs-lookup"><span data-stu-id="7620a-194">Link opens in add-in task pane.</span></span>|
|<span data-ttu-id="7620a-195">Windows 版 Outlook 2016（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="7620a-195">Outlook 2016 on Windows (one-time purchase)</span></span><br><span data-ttu-id="7620a-196">Windows 上的 Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="7620a-196">Outlook 2013 on Windows</span></span>|<span data-ttu-id="7620a-197">否</span><span class="sxs-lookup"><span data-stu-id="7620a-197">No</span></span>|<span data-ttu-id="7620a-198">链接将在 Internet Explorer 11 中打开。</span><span class="sxs-lookup"><span data-stu-id="7620a-198">Link opens in Internet Explorer 11.</span></span>|
|<span data-ttu-id="7620a-199">其他客户端</span><span class="sxs-lookup"><span data-stu-id="7620a-199">Other clients</span></span>|<span data-ttu-id="7620a-200">否</span><span class="sxs-lookup"><span data-stu-id="7620a-200">No</span></span>|<span data-ttu-id="7620a-201">链接将在用户的默认浏览器中打开。</span><span class="sxs-lookup"><span data-stu-id="7620a-201">Link opens in user's default browser.</span></span>|

<span data-ttu-id="7620a-202">有关更多详细信息，请参阅[指定要在加载项窗口中打开的域](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window)。</span><span class="sxs-lookup"><span data-stu-id="7620a-202">For more details, see the [Specify domains you want to open in the add-in window](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).</span></span>

## <a name="permissions"></a><span data-ttu-id="7620a-203">权限</span><span class="sxs-lookup"><span data-stu-id="7620a-203">Permissions</span></span>

<span data-ttu-id="7620a-p119">**Permissions** 元素包含外接程序所需的权限。通常情况下，应指定外接程序所需的最低权限，具体视计划要使用的确切方法而定。例如，如果在撰写窗体中激活的邮件外接程序对 [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 等项属性只执行读取操作，而不执行写入操作，也不调用 [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 访问任何 Exchange Web 服务操作，应指定 **ReadItem** 权限。若要详细了解可用权限，请参阅 [了解 Outlook 外接程序权限](understanding-outlook-add-in-permissions.md)。</span><span class="sxs-lookup"><span data-stu-id="7620a-p119">The **Permissions** element contains the required permissions for the add-in. In general, you should specify the minimum necessary permission that your add-in needs, depending on the exact methods that you plan to use. For example, a mail add-in that activates in compose forms and only reads but does not write to item properties like [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), and does not call [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) to access any Exchange Web Services operations should specify **ReadItem** permission. For details on the available permissions, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

<span data-ttu-id="7620a-208">**邮件外接程序的 4 层权限模型**</span><span class="sxs-lookup"><span data-stu-id="7620a-208">**Four-tier permissions model for mail add-ins**</span></span>

![邮件应用架构 v1.1 的 4 层权限模型。](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## <a name="activation-rules"></a><span data-ttu-id="7620a-210">激活规则</span><span class="sxs-lookup"><span data-stu-id="7620a-210">Activation rules</span></span>

<span data-ttu-id="7620a-p120">**Rule** 元素中指定了激活规则。**Rule** 元素可以显示为 1.1 清单中的 **OfficeApp** 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="7620a-p120">Activation rules are specified in the **Rule** element. The **Rule** element can appear as a child of the **OfficeApp** element in 1.1 manifests.</span></span>

<span data-ttu-id="7620a-213">激活规则可用于根据当前所选项目的下列一个或多个条件激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="7620a-213">Activation rules can be used to activate an add-in based on one or more of the following conditions on the currently selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="7620a-214">激活规则只适用于不支持 **VersionOverrides** 元素的客户端。</span><span class="sxs-lookup"><span data-stu-id="7620a-214">Activation rules only apply to clients that do not support the **VersionOverrides** element.</span></span>

- <span data-ttu-id="7620a-215">项目类型和/或邮件类别</span><span class="sxs-lookup"><span data-stu-id="7620a-215">The item type and/or message class</span></span>

- <span data-ttu-id="7620a-216">存在特定类型的已知实体，例如地址或电话号码</span><span class="sxs-lookup"><span data-stu-id="7620a-216">The presence of a specific type of known entity, such as an address or phone number</span></span>

- <span data-ttu-id="7620a-217">正文、主题或发件人电子邮件地址中的正则表达式匹配</span><span class="sxs-lookup"><span data-stu-id="7620a-217">A regular expression match in the body, subject, or sender email address</span></span>

- <span data-ttu-id="7620a-218">存在附件</span><span class="sxs-lookup"><span data-stu-id="7620a-218">The presence of an attachment</span></span>

<span data-ttu-id="7620a-219">有关激活规则的详细信息和示例，请参阅 [Outlook 外接程序的激活规则](activation-rules.md)。</span><span class="sxs-lookup"><span data-stu-id="7620a-219">For details and samples of activation rules, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>


## <a name="next-steps-add-in-commands"></a><span data-ttu-id="7620a-220">后续步骤：外接程序命令</span><span class="sxs-lookup"><span data-stu-id="7620a-220">Next steps: Add-in commands</span></span>

<span data-ttu-id="7620a-p121">定义基本清单后， 为外接程序定义外接程序命令。外接程序命令代表功能区中的按钮，因此用户以一种简单、直观的方式激活外接程序。有关详细信息，请参阅[用于 Outlook 的外接程序命令](add-in-commands-for-outlook.md)。</span><span class="sxs-lookup"><span data-stu-id="7620a-p121">After defining a basic manifest, define add-in commands for your add-in. Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="7620a-224">有关定义外接程序命令的示例外接程序，请参阅 [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo)。</span><span class="sxs-lookup"><span data-stu-id="7620a-224">For an example add-in that defines add-in commands, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).</span></span>

## <a name="next-steps-add-mobile-support"></a><span data-ttu-id="7620a-225">后续步骤：添加移动支持</span><span class="sxs-lookup"><span data-stu-id="7620a-225">Next steps: Add mobile support</span></span>

<span data-ttu-id="7620a-p122">外接程序可选择性的为 Outlook Mobile 添加支持。Outlook Mobile 支持外接程序命令的方式与在 Windows 和 Mac 上使用 Outlook 的方式类似。有关详细信息，请参阅[为 Outlook Mobile 的外接程序命令添加支持](add-mobile-support.md)</span><span class="sxs-lookup"><span data-stu-id="7620a-p122">Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and Mac. For more information, see [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="7620a-229">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7620a-229">See also</span></span>

- [<span data-ttu-id="7620a-230">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="7620a-230">Localization for Office Add-ins</span></span>](../develop/localization.md)
- [<span data-ttu-id="7620a-231">Outlook 外接程序的隐私、权限和安全性</span><span class="sxs-lookup"><span data-stu-id="7620a-231">Privacy, permissions, and security for Outlook add-ins</span></span>](privacy-and-security.md)
- [<span data-ttu-id="7620a-232">Outlook 外接程序 API</span><span class="sxs-lookup"><span data-stu-id="7620a-232">Outlook add-in APIs</span></span>](apis.md)
- [<span data-ttu-id="7620a-233">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="7620a-233">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="7620a-234">Office 外接程序清单的架构参考 (v1.1)</span><span class="sxs-lookup"><span data-stu-id="7620a-234">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="7620a-235">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="7620a-235">Design your Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="7620a-236">了解 Outlook 外接程序权限</span><span class="sxs-lookup"><span data-stu-id="7620a-236">Understanding Outlook add-in permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="7620a-237">使用正则表达式激活规则显示 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="7620a-237">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="7620a-238">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="7620a-238">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
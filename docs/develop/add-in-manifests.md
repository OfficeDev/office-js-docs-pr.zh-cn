---
title: Office 加载项 XML 清单
description: 获取 Office 加载项清单及其用途概述。
ms.date: 05/24/2022
ms.localizationpriority: high
ms.openlocfilehash: 09b4d5b2b9fc92c977217df94730b3e6e56cacaa
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659988"
---
# <a name="office-add-ins-xml-manifest"></a>Office 加载项 XML 清单

Office 外接程序的 XML 清单文件描述，当最终用户安装外接程序并将其与 Office 文档和应用程序配合使用时，应如何激活外接程序。

> [!TIP]
> 本文介绍当前的 XML 格式清单。 还有一个 JSON 格式的 Teams 清单以预览版提供。 有关详细信息，请参阅 [Office 加载项的 Teams 清单（预览版）](json-manifest-overview.md)。

XML 清单文件支持 Office 加载项执行以下操作：

- 通过提供 ID、版本、说明、显示名称和默认区域设置进行自我描述。

- 指定用于为加载项塑造品牌的图像，以及用于 Office 应用功能区中[加载项命令](create-addin-commands.md)的图标。

- 指定外接程序如何与 Office 集成，包括任何自定义 UI，如外接程序创建的功能区按钮。

- 指定内容外接程序请求的默认尺寸和 Outlook 外接程序请求的高度。

- 声明 Office 外接程序所需的权限，例如读取或写入文档。

- 对于 Outlook 外接程序，定义一个或多个规则，以指定将在其中激活规则并与邮件、约会或会议请求项目交互的上下文。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="required-elements"></a>必需元素

下表指定了三种类型 Office 加载项的必需元素。

> [!NOTE]
> 还存在强制性命令，其中元素必须出现在其父元素中。 有关详细信息，请参阅[如何查找清单元素的正确顺序](manifest-element-ordering.md)。

### <a name="required-elements-by-office-add-in-type"></a>Office 加载项类型的必需元素

| 元素                                                                                      | 内容    | 任务窗格    | Outlook      |
| :------------------------------------------------------------------------------------------- | :--------: | :----------: | :--------:   |
| [OfficeApp][]                                                                                | 必需   | 必需     | 必需     |
| [Id][]                                                                                       | 必需   | 必需     | 必需     |
| [版本][]                                                                                  | 必需   | 必需     | 必需     |
| [ProviderName][]                                                                             | 必需   | 必需     | 必需     |
| [DefaultLocale][]                                                                            | 必需   | 必需     | 必需     |
| [DisplayName][]                                                                              | 必需   | 必需     | 必需     |
| [Description][]                                                                              | 必需   | 必需     | 必需     |
| [IconUrl][]                                                                                  | 必需   | 必需     | 必需     |
| [SupportUrl][]\*\*                                                                           | 必需   | 必需     | 必需     |
| [DefaultSettings (ContentApp)][]<br/>[DefaultSettings (TaskPaneApp)][]                       | 必需   | 必需     | 不可用|
| [SourceLocation (ContentApp)][]<br/>[SourceLocation (TaskPaneApp)][]<br/>[SourceLocation (MailApp)][]| 必需 | 必需 | 必需   |
| [DesktopSettings][]                                                                          | 不可用 | 不可用 | 必需 |
| [Permissions (ContentApp)][]<br/>[Permissions (TaskPaneApp)][]<br/>[Permissions (MailApp)][] | 必需   | 必需     | 必需     |
| [Rule (RuleCollection)][]<br/>[Rule (MailApp)][]                                             | 不可用 | 不可用 | 必需 |
| [要求 （MailApp）][]\*                                                                 | 不适用| 不可用 | 必需 |
| [设置][]\*<br/>[集（要求）][]\*<br/>[集 （MailAppRequirements）][]\*                 | 必需   | 必需     | 必需     |
| [表单][]\*<br/>[FormSettings][]\*                                                            | 不可用 | 不可用 | 必需 |
| [主机][]\*                                                                                  | 必需   | 必需     | 可选     |

_\*Office 加载项清单架构版本 1.1 中新增_

_\*\* 仅通过 AppSource 分发的加载项才需要 SupportUrl。_

<!-- Links for above table -->

[officeapp]: /javascript/api/manifest/officeapp
[id]: /javascript/api/manifest/id
[version]: /javascript/api/manifest/version
[providername]: /javascript/api/manifest/providername
[defaultlocale]: /javascript/api/manifest/defaultlocale
[displayname]: /javascript/api/manifest/displayname
[description]: /javascript/api/manifest/description
[iconurl]: /javascript/api/manifest/iconurl
[supporturl]: /javascript/api/manifest/supporturl
[defaultsettings (contentapp)]: /javascript/api/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: /javascript/api/manifest/defaultsettings
[sourcelocation (contentapp)]: /javascript/api/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: /javascript/api/manifest/sourcelocation
[sourcelocation (mailapp)]: /javascript/api/manifest/sourcelocation
[desktopsettings]: /javascript/api/manifest/desktopsettings
[permissions (contentapp)]: /javascript/api/manifest/permissions
[permissions (taskpaneapp)]: /javascript/api/manifest/permissions
[permissions (mailapp)]: /javascript/api/manifest/permissions
[rule (rulecollection)]: /javascript/api/manifest/rule
[rule (mailapp)]: /javascript/api/manifest/rule
[要求 （mailapp）]: /javascript/api/manifest/requirements
[set]: /javascript/api/manifest/set
[集 （mailapprequirements）]: /javascript/api/manifest/sets
[表单]: /javascript/api/manifest/form
[formsettings]: /javascript/api/manifest/formsettings
[集（要求）]: /javascript/api/manifest/sets
[主机]: /javascript/api/manifest/hosts

## <a name="hosting-requirements"></a>托管要求

所有图像 URI（如用于[外接程序命令](create-addin-commands.md)的 URI）都必须支持缓存。 托管图像的服务器不得在 HTTP 响应中返回指定 `no-cache`、`no-store` 或类似选项的 `Cache-Control` 标头。

所有 URL（如 [SourceLocation](/javascript/api/manifest/sourcelocation) 元素中指定的源文件位置）都应 **受 SSL 保护 (HTTPS)**。 [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a>关于提交到 AppSource 的最佳做法

确保外接程序 ID 有效且具有唯一 GUID。Web 上提供可用于创建唯一 GUID 的各种 GUID 生成器工具。

提交到 AppSource 的加载项还必须包括 [SupportUrl](/javascript/api/manifest/supporturl) 元素。 有关详细信息，请参阅[提交到 AppSource 的应用和加载项的验证策略](/legal/marketplace/certification-policies)。

仅使用 [AppDomain](/javascript/api/manifest/appdomains) 元素指定域（除了在 [SourceLocation](/javascript/api/manifest/sourcelocation) 元素中指定的用于身份验证方案的域）。

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a>指定要在外接程序窗口中打开的域

当在 web 上的 Office 中运行时，任务窗格可以导航到任何 URL。但在桌面平台中，如果加载项尝试转到托管起始页（如清单文件的 [SourceLocation](/javascript/api/manifest/sourcelocation) 元素中所指定的）的域之外的域中的 URL，则该 URL 将在 Office 应用程序的加载项窗格外的新浏览器窗口中打开。

若要重写此（桌面版 Office）操作，请在清单文件的 [AppDomains](/javascript/api/manifest/appdomains) 元素中指定的域列表中指定要在外接程序窗口中打开的每个域。 如果加载项尝试转至该列表的域中的 URL，则它将在 Office 网页版和桌面版中的任务窗口中打开。 如果它尝试转至列表之外的域中的 URL，则在桌面版 Office 中，该 URL 将在新的浏览器窗口中（外接程序窗格之外）打开。

> [!NOTE]
> 该行为有两个例外情况。
>
> - 它仅适用于外接程序的根窗格。 如果外接程序页面中嵌入有 iframe，则可以将该 iframe 定向到任何 URL，不论它是否列在 **\<AppDomains\>** 中，即使在桌面版 Office 中也是如此。
> - 使用 [displayDialogAsync](/javascript/api/office/office.ui?view=common-js&preserve-view=true#office-office-ui-displaydialogasync-member(1)) API 打开对话框时，传递到方法的 URL 必须与外接程序位于相同的域，但是之后对话框可以定向到任意 URL，无论其是否列入 **\<AppDomains\>** 甚至桌面 Office 中。

以下 XML 清单示例在 **\<SourceLocation\>** 元素中指定的 `https://www.contoso.com` 域中托管其外接程序页面。 它还指定 **\<AppDomains\>** 元素列表内 [AppDomain](/javascript/api/manifest/appdomain) 元素中的 `https://www.northwindtraders.com` 域。 如果加载项转到 `www.northwindtraders.com` 域中的页面，此页面会在加载项窗格中打开，即使是在 Office 桌面版中，也不例外。

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="version-overrides-in-the-manifest"></a>清单中的版本替代

可选的 [VersionOverrides](/javascript/api/manifest/versionoverrides) 元素值得特别提及。 它包含支持其他加载项功能的子标记。 其中一些为：

- 自定义 Office 功能区和菜单。
- 自定义 Office 与加载项在其中运行的嵌入式浏览器运行时一起工作的方式。
- 配置加载项如何与 Azure Active Directory 和 Microsoft Graph 交互以进行单一登录。

`VersionOverrides` 的一些子代元素具有替代父级 `OfficeApp` 元素值的值。 例如，`VersionOverrides` 中的 `Hosts` 元素替代 `OfficeApp` 中的 `Hosts` 元素。

`VersionOverrides` 元素具有其自己的架构（实际上有四个架构），具体取决于加载项的类型及其使用的功能。这些架构是：

- [任务窗格 1.0](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)
- [内容 1.0](/openspecs/office_file_formats/ms-owemxml/c9cb8dca-e9e7-45a7-86b7-f1f0833ce2c7)
- [邮件 1.0](/openspecs/office_file_formats/ms-owemxml/578d8214-2657-4e6a-8485-25899e772fac)
- [邮件 1.1](/openspecs/office_file_formats/ms-owemxml/8e722c85-eb78-438c-94a4-edac7e9c533a)

在使用 `VersionOverrides` 元素时，`OfficeApp` 元素必须具有标识相应架构的 `xmlns` 属性。 属性的可能值如下：

- `http://schemas.microsoft.com/office/taskpaneappversionoverrides`
- `http://schemas.microsoft.com/office/contentappversionoverrides`
- `http://schemas.microsoft.com/office/mailappversionoverrides`

`VersionOverrides` 元素本身还必须具有 `xmlns` 属性来指定架构。 可能的值包括上述三个和以下值：

- `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`

`VersionOverrides`元素还必须具有指定架构版本的 `xsi:type` 属性。 可能的值如下：

- `VersionOverridesV1_0`
- `VersionOverridesV1_1`

以下是在任务窗格加载项和邮件加载项中分别使用的 `VersionOverrides` 的示例。 请注意，在使用版本 1.1 的邮件 `VersionOverrides` 时，它必须是类型 1.0 的父级 `VersionOverrides` 的最后一个子级。 内部 `VersionOverrides` 中子元素的值替代父级 `VersionOverrides` 和祖父级 `OfficeApp` 元素中同名元素的值。

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <!-- child elements omitted -->
</VersionOverrides>
```

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <!-- other child elements omitted -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <!-- child elements omitted -->
  </VersionOverrides>
</VersionOverrides>
```

有关包含 `VersionOverrides` 元素的清单示例，请参阅 [清单 v1.1 XML 文件示例和架构](#manifest-v11-xml-file-examples-and-schemas)。

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a>指定从中执行 Office .js API 调用的域

你的加载项可以从清单文件的 [SourceLocation](/javascript/api/manifest/sourcelocation) 元素中引用的域执行 Office.js API 调用。 如果加载项中有需要访问 Office.js API 的其他 IFrame，请将该源 URL 的域添加到在清单文件的 [AppDomains](/javascript/api/manifest/appdomains) 元素中指定的列表。 如果有一个未包含在 `AppDomains` 列表中且具有源的 IFrame 尝试执行 Office.js API 调用，则加载项将收到[“权限被拒绝”错误](../reference/javascript-api-for-office-error-codes.md)。

## <a name="manifest-v11-xml-file-examples-and-schemas"></a>清单 v1.1 XML 文件示例和架构

下面各部分展示了内容加载项、任务窗格加载项和 Outlook 加载项的清单 v1.1 XML 文件示例。

# <a name="task-pane"></a>[任务窗格](#tab/tabid-1)

[加载项清单架构](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
  <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office app ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
              <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                  <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                  <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                    <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                    <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon80" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
                    <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
                <!-- Menu example -->
                <Control xsi:type="Menu" id="Contoso.Menu">
                  <Label resid="Contoso.Dropdown.Label" />
                  <Supertip>
                    <Title resid="Contoso.Dropdown.Label" />
                    <Description resid="Contoso.Dropdown.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane1.Url" />
                      </Action>
                    </Item>

                    <Item id="Contoso.Menu.Item2">
                      <Label resid="Contoso.Item2.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item2.Label" />
                        <Description resid="Contoso.Item2.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane2.Url" />
                      </Action>
                    </Item>

                  </Items>
                </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon16" DefaultValue="https://myCDN/Images/Button16x16.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon32" DefaultValue="https://myCDN/Images/Button32x32.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon80" DefaultValue="https://myCDN/Images/Button80x80.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://myCDN/Images/ButtonFunction.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# <a name="content"></a>[内容](#tab/tabid-2)

[加载项清单架构](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# <a name="mail"></a>[邮件](#tab/tabid-3)

[加载项清单架构](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## <a name="validate-an-office-add-ins-manifest"></a>验证 Office 加载项的清单

有关根据 [XML 架构定义 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) 验证清单的信息，请参阅[验证 Office 加载项的清单](../testing/troubleshoot-manifest.md)。

## <a name="see-also"></a>另请参阅

- [如何查找清单元素的正确顺序](manifest-element-ordering.md)
- [在清单中创建外接程序命令](create-addin-commands.md)
- [指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md)
- [Office 外接程序的本地化](localization.md)
- [Office 外接程序清单的架构参考](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
- [更新 API 和清单版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)
- [标识等效的 COM 加载项](make-office-add-in-compatible-with-existing-com-add-in.md)
- [在加载项中请求获取 API 使用权限](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
- [验证 Office 加载项的清单](../testing/troubleshoot-manifest.md)

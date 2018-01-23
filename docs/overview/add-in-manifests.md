
# <a name="office-add-ins-xml-manifest"></a>Office 外接程序 XML 清单


Office 外接程序的 XML 清单文件描述，当最终用户安装外接程序并将其与 Office 文档和应用程序配合使用时，应如何激活外接程序。 

基于此架构的 XML 清单文件允许 Office 外接程序执行以下内容：

- 通过提供 ID、版本、描述、显示名称和默认区域设置进行自我描述。
    
- 指定如何将外接程序与 Office 集成，其中包括任何自定义的 UI，例如，外接程序创建的功能区按钮。
    
- 指定内容外接程序请求的默认尺寸和 Outlook 外接程序请求的高度。
    
- 声明 Office 外接程序所需的权限，例如读取或写入文档。
    
- 对于 Outlook 外接程序，定义一个或多个规则，以指定将在其中激活规则并与邮件、约会或会议请求项目交互的上下文。

>
  **注意：**生成外接程序时，如果计划将外接程序[发布](../publish/publish.md)到 Office 应用商店，请务必遵循 [Office 应用商店验证策略](https://msdn.microsoft.com/en-us/library/jj220035.aspx)。例如，外接程序必须适用于支持你定义的方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3)以及 [Office 外接程序主机和可用性](https://dev.office.com/add-in-availability)页）。


## <a name="required-elements"></a>必需元素


下表指定三种类型的 Office 外接程序必需的元素。


**Office 外接程序类型的必需元素**


|**元素**|**内容**|**任务窗格**|**Outlook**|
|:-----|:-----|:-----|:-----|
|[OfficeApp](http://msdn.microsoft.com/en-us/library/68f1cada-66f8-4341-45f5-14e0634c24fb%28Office.15%29.aspx)|X|X|X|
|[Id](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx)|X|X|X|
|[版本](http://msdn.microsoft.com/en-us/library/6a8bbaa5-ee8c-6824-4aba-cb1a804269f6%28Office.15%29.aspx)|X|X|X|
|[ProviderName](http://msdn.microsoft.com/en-us/library/0062693a-fafa-ea2d-051a-75dac0f6c323%28Office.15%29.aspx)|X|X|X|
|[DefaultLocale](http://msdn.microsoft.com/en-us/library/04796a3a-3afa-dc85-db66-4677560c185c%28Office.15%29.aspx)|X|X|X|
|[DisplayName](http://msdn.microsoft.com/en-us/library/529159ca-53bf-efcf-c245-e572dab0ef57%28Office.15%29.aspx)|X|X|X|
|[说明](http://msdn.microsoft.com/en-us/library/bcce6bad-23d0-7631-7d8c-1064b8453b5a%28Office.15%29.aspx)|X|X|X|
|[IconUrl](http://msdn.microsoft.com/library/c7dac2d4-4fda-6fc7-3774-49f02b2d3e1e%28Office.15%29.aspx)|X|X|X|
|[HighResolutionIconUrl](http://msdn.microsoft.com/library/ff7b2647-ec8e-70dc-4e4a-e1a1377ff3f2%28Office.15%29.aspx)|X|X|X|
|[DefaultSettings (ContentApp)](http://msdn.microsoft.com/en-us/library/f7edc689-551f-1a17-ea81-ffd58f534557%28Office.15%29.aspx)<br/>[DefaultSettings (TaskPaneApp)](http://msdn.microsoft.com/en-us/library/36e3d139-56a4-fb3d-0a21-cbd14e606765%28Office.15%29.aspx)|X|X||
|[SourceLocation (ContentApp)](http://msdn.microsoft.com/en-us/library/00d95bb0-e8f5-647f-790a-0aa3aabc8141%28Office.15%29.aspx)<br/>[SourceLocation (TaskPaneApp)](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx)|X|X||
|[DesktopSettings](http://msdn.microsoft.com/en-us/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx)|||X|
|[SourceLocation (MailApp)](http://msdn.microsoft.com/en-us/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx)|||X|
|[Permissions (ContentApp)](http://msdn.microsoft.com/en-us/library/9f3dcf9c-fced-c115-4f0d-38d60fb7c583%28Office.15%29.aspx)<br/>[Permissions (TaskPaneApp)](http://msdn.microsoft.com/en-us/library/d4cfe645-353d-8240-8495-f76fb36602fe%28Office.15%29.aspx)<br/>[Permissions (MailApp)](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx)|X|X|X|
|[Rule (RuleCollection)](http://msdn.microsoft.com/en-us/library/c6ce9d52-4b53-c6a6-de7e-c64106135c81%28Office.15%29.aspx)<br/>[Rule (MailApp)](http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx)|||X|
|[*Requirements (MailApp)](http://msdn.microsoft.com/en-us/library/9536ea30-34f7-76b5-7f30-1508626840e4%28Office.15%29.aspx)|||X|
|[*Set](http://msdn.microsoft.com/en-us/library/1506daa1-332c-30e1-6402-3371bcd0b895%28Office.15%29.aspx)<br/>[**Sets (MailAppRequirements)](http://msdn.microsoft.com/en-us/library/2a6a2484-eeee-37e4-43bc-c185e8ae0d1d%28Office.15%29.aspx)|||X|
|[*Form](http://msdn.microsoft.com/en-us/library/77a8ac83-c22b-1225-4fc4-ba4038b68648%28Office.15%29.aspx)<br/>[**FormSettings](http://msdn.microsoft.com/en-us/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx)|||X|
|[*Sets (Requirements)](http://msdn.microsoft.com/en-us/library/509be287-b532-87c6-71ac-64f3a4bbd3af%28Office.15%29.aspx)|X|X||
|[*Hosts](http://msdn.microsoft.com/library/f9a739c1-3daf-c03a-2bd9-4a2a6b870101%28Office.15%29.aspx)|X|X||

*已添加在 Office 外接程序清单架构版本 1.1 中。

**重要说明**： 

- 确保外接程序 ID 有效且具有唯一 GUID。Web 上提供可用于创建唯一 GUID 的各种 GUID 生成器工具。 

- 所有 URL（例如 [SourceLocation](http://dev.office.com/reference/add-ins/manifest/sourcelocation) 元素中指定的源文件位置）必须采用 **SSL 保护 (HTTPS)**。

- 所有图标 URL（例如在命令界面上使用的 URL）必须**允许缓存**。Web 服务器不应返回类似不缓存/不存储的 HTTP 头。 

- 提交到 Office 应用商店的外接程序还必须包括 [SupportUrl](http://dev.office.com/reference/add-ins/manifest/supporturl) 元素。有关详细信息，请参阅 [要避免的一些常见提交错误是什么？](http://msdn.microsoft.com/library/0ceb385c-a608-40cc-8314-78e39d6c75d0%28Office.15%29.aspx#bk_q2)

- 作为一种最佳做法，仅使用 [AppDomain](http://dev.office.com/reference/add-ins/manifest/appdomains) 元素指定身份验证方案除在 [SourceLocation](http://dev.office.com/reference/add-ins/manifest/sourcelocation) 元素中指定的域之外的域。

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a>指定想要在外接程序窗口中打开的域

默认情况下，如果外接程序尝试转到托管起始页（如清单文件的 [SourceLocation](http://dev.office.com/reference/add-ins/manifest/sourcelocation) 元素中所指定的）的域之外的域中的 URL，则该 URL 将在 Office 主机应用程序的外接程序窗格外的新浏览器窗口中打开。此默认行为将防止用户在外接程序窗格从嵌入的 **iframe** 元素中进行意外页面导航。

若要重写此操作，请在清单文件的 [AppDomains](http://dev.office.com/reference/add-ins/manifest/appdomains) 元素中指定的域列表中指定要在外接程序窗口中打开的每个域。如果外接程序尝试转至在列表之外的域的 URL，则该 URL 将在新的浏览器窗口中（外接程序窗格之外）打开。

以下 XML 清单示例在 **SourceLocation** 元素中指定的 `https://www.contoso.com` 域中托管其外接程序主页。它还指定 [AppDomain](http://msdn.microsoft.com/en-us/library/2a0353ec-5e09-6fbf-1636-4bb5dcebb9bf%28Office.15%29.aspx) 元素列表内的 **AppDomains** 元素中的 `https://www.northwindtraders.com` 域。如果外接程序转至 www.northwindtraders.com 域中的页面，则在外接程序窗格中打开该页面。


```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="manifest-v11-xml-file-examples-and-schemas"></a>清单 v1.1 XML 文件示例和架构


下面各节显示了内容、任务窗格和 Outlook 外接程序的清单 v1.1 XML 文件的示例。

### <a name="office-add-in-manifest-v11-example-with-commands-and-fallback-task-pane"></a>带有命令和回退任务窗格的 Office 外接程序清单 v1.1 示例
[任务窗格清单架构](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas/taskpane)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

<!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

<!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

<!--IMPORTANT! Id must be unique for your add-in. If you clone this manifest ensure that you change this id to your own GUID -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
   <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
   <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://i.imgur.com/oZFS95h.png" />

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
        <!-- Form factor. Currenly only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
            <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called. 
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office Ribbon-->
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
                <!--Icons. Required sizes 16,31,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX -->
                <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFuncion or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  - <Action xsi:type="ExecuteFunction">
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
              </Icon>
              <Items>
                <Item id="Contoso.Menu.Item1">
                  <Label resid="Contoso.Item1.Label"/>
                  <Supertip>
                    <Title resid="Contoso.Item1.Label" />
                    <Description resid="Contoso.Item1.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
              <!-- If validating with XSD it needs to be at the end, we might change this before release -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon" DefaultValue="https://i.imgur.com/FkSShX9.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://i.imgur.com/qDujiX0.png" />
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

### <a name="content-add-in-manifest-v11-example"></a>内容外接程序清单 v1.1 示例
[内容清单架构](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas/content)


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/ENUSIcon.png" />
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

### <a name="outlook-add-in-manifest-v11-example"></a>Outlook 外接程序清单 v1.1 示例
[内容清单架构](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas/mail)


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="MailApp">

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
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following line to specify    -->
  <!-- the web serverthat hosts the icon file. -->
  <IconUrl DefaultValue=
    "https://webserver/YouTube/YouTubeLogo.png"/>

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

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>验证并排查清单问题

如需排查清单问题，请参阅[验证并排查清单问题](../testing/troubleshoot-manifest.md)。其中介绍了如何针对 [XML 架构定义 (XSD)](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas) 验证清单，以及如何使用运行时日志记录功能调试清单。

## <a name="additional-resources"></a>其他资源


- [在清单中定义外接程序命令](../develop/define-add-in-commands.md)
- [指定 Office 主机和 API 要求](../overview/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序的本地化](../develop/localization.md)
- [Office 外接程序清单的架构参考](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas)
- [验证并排查清单问题](../testing/troubleshoot-manifest.md)


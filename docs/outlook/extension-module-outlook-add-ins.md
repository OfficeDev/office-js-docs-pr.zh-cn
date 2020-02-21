---
title: 模块扩展 Outlook 加载项
description: 可以创建在 Outlook 中运行的应用程序，以便用户无需退出 Outlook 即可轻松地访问业务信息和工作效率工具。
ms.date: 06/04/2019
localization_priority: Normal
ms.openlocfilehash: a7c4031061dd5cd4605220c6839481b3a88e53d9
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166018"
---
# <a name="module-extension-outlook-add-ins"></a><span data-ttu-id="3191a-103">模块扩展 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="3191a-103">Module extension Outlook add-ins</span></span>

<span data-ttu-id="3191a-104">模块扩展加载项与邮件、任务和日历一起显示在 Outlook 导航栏中。</span><span class="sxs-lookup"><span data-stu-id="3191a-104">Module extension add-ins appear in the Outlook navigation bar, right alongside mail, tasks, and calendars.</span></span> <span data-ttu-id="3191a-105">模块扩展不限于使用邮件和约会信息。</span><span class="sxs-lookup"><span data-stu-id="3191a-105">A module extension is not limited to using mail and appointment information.</span></span> <span data-ttu-id="3191a-106">可以创建在 Outlook 中运行的应用程序，以便用户无需退出 Outlook 即可轻松地访问业务信息和工作效率工具。</span><span class="sxs-lookup"><span data-stu-id="3191a-106">You can create applications that run inside Outlook to make it easy for your users to access business information and productivity tools without ever leaving Outlook.</span></span>

> [!NOTE]
> <span data-ttu-id="3191a-107">仅 Windows 上的 Outlook 2016 或更高版本支持模块扩展。</span><span class="sxs-lookup"><span data-stu-id="3191a-107">Module extensions are only supported by Outlook 2016 or later on Windows.</span></span>  

## <a name="open-a-module-extension"></a><span data-ttu-id="3191a-108">打开模块扩展</span><span class="sxs-lookup"><span data-stu-id="3191a-108">Open a module extension</span></span>

<span data-ttu-id="3191a-p102">要打开模块扩展，用户单击 Outlook 导航栏中的模块的名称或图标即可。如果用户选择了紧凑型导航，导航栏有一个显示已加载扩展的图标。</span><span class="sxs-lookup"><span data-stu-id="3191a-p102">To open a module extension, users click on the module's name or icon in the Outlook navigation bar. If the user has compact navigation selected, the navigation bar has an icon that shows an extension is loaded.</span></span>

![当模块扩展在 Outlook 中加载时，显示紧凑型导航栏。](../images/outlook-module-navigationbar-compact.png)

<span data-ttu-id="3191a-112">如果用户没有使用紧凑型导航，则导航栏有两种外观。</span><span class="sxs-lookup"><span data-stu-id="3191a-112">If the user is not using compact navigation, the navigation bar has two looks.</span></span> <span data-ttu-id="3191a-113">加载一个扩展后，它将显示加载项的名称。</span><span class="sxs-lookup"><span data-stu-id="3191a-113">With one extension loaded, it shows the name of the add-in.</span></span>

![当一个模块扩展在 Outlook 中加载时，显示展开的导航栏。](../images/outlook-module-navigationbar-one.png)

<span data-ttu-id="3191a-115">在加载多个加载项时，会显示**加载项**一词。单击其中任何一个即可打开扩展的用户界面。</span><span class="sxs-lookup"><span data-stu-id="3191a-115">When more than one add-in is loaded, it shows the word **Add-ins**. Clicking either will open the extension's user interface.</span></span>

![当多个模块扩展在 Outlook 中加载时，显示展开的导航栏。](../images/outlook-module-navigationbar-more.png)

<span data-ttu-id="3191a-117">在单击扩展时，Outlook 会将内置模块替换为自定义模块，以便用户可以与该加载项进行交互。</span><span class="sxs-lookup"><span data-stu-id="3191a-117">When you click on an extension, Outlook replaces the built-in module with your custom module so that your users can interact with the add-in.</span></span> <span data-ttu-id="3191a-118">你可以使用外接程序中 Outlook JavaScript API 的所有功能，可以在与外接程序内容交互的 Outlook 功能区中创建命令按钮。</span><span class="sxs-lookup"><span data-stu-id="3191a-118">You can use all of the features of the Outlook JavaScript API in your add-in, and can create command buttons in the Outlook ribbon that will interact with the add-in content.</span></span> <span data-ttu-id="3191a-119">以下屏幕截图显示集成在 Outlook 导航栏中的加载项，并拥有将更新该加载项内容的功能区命令。</span><span class="sxs-lookup"><span data-stu-id="3191a-119">The following screenshot shows an add-in that is integrated in the Outlook navigation bar and has ribbon commands that will update the content of the add-in.</span></span>

![显示模块扩展的用户界面](../images/outlook-module-extension.png)

## <a name="example"></a><span data-ttu-id="3191a-121">示例</span><span class="sxs-lookup"><span data-stu-id="3191a-121">Example</span></span>

<span data-ttu-id="3191a-122">下面是定义模块扩展的清单文件部分。</span><span class="sxs-lookup"><span data-stu-id="3191a-122">The following is a section of a manifest file that defines a module extension.</span></span>

```xml
<!-- Add Outlook module extension point -->
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"
                  xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1"
                    xsi:type="VersionOverridesV1_1">

    <!-- Begin override of existing elements -->
    <Description resid="residVersionOverrideDesc" />

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <!-- End override of existing elements -->

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <!-- Set the URL of the file that contains the
                JavaScript function that controls the extension -->
          <FunctionFile resid="residFunctionFileUrl" />

          <!--New Extension Point - Module for a ModuleApp -->
          <ExtensionPoint xsi:type="Module">
            <SourceLocation resid="residExtensionPointUrl" />
            <Label resid="residExtensionPointLabel" />

            <CommandSurface>
              <CustomTab id="idTab">
                <Group id="idGroup">
                  <Label resid="residGroupLabel" />

                  <Control xsi:type="Button" id="group.changeToAssociate">
                    <Label resid="residChangeToAssociateLabel" />
                    <Supertip>
                      <Title resid="residChangeToAssociateLabel" />
                      <Description resid="residChangeToAssociateDesc" />
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="residAssociateIcon16" />
                      <bt:Image size="32" resid="residAssociateIcon32" />
                      <bt:Image size="80" resid="residAssociateIcon80" />
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>changeToAssociateRate</FunctionName>
                    </Action>
                  </Control>
                  
              </Group>
                <Label resid="residCustomTabLabel" />
              </CustomTab>
            </CommandSurface>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="residAddinIcon16" 
                  DefaultValue="https://localhost:8080/Executive-16.png" />
        <bt:Image id="residAddinIcon32" 
                  DefaultValue="https://localhost:8080/Executive-32.png" />
        <bt:Image id="residAddinIcon80" 
                  DefaultValue="https://localhost:8080/Executive-80.png" />
      
        <bt:Image id="residAssociateIcon16" 
                  DefaultValue="https://localhost:8080/Associate-16.png" />
        <bt:Image id="residAssociateIcon32" 
                  DefaultValue="https://localhost:8080/Associate-32.png" />
        <bt:Image id="residAssociateIcon80" 
                  DefaultValue="https://localhost:8080/Associate-80.png" />
      </bt:Images>

      <bt:Urls>
        <bt:Url id="residFunctionFileUrl" 
                DefaultValue="https://localhost:8080/" />
        <bt:Url id="residExtensionPointUrl" 
                DefaultValue="https://localhost:8080/" />
      </bt:Urls>

      <!--Short strings must be less than 30 characters long -->
      <bt:ShortStrings>
        <bt:String id="residExtensionPointLabel" 
                    DefaultValue="Billable Hours" />
        <bt:String id="residGroupLabel" 
                    DefaultValue="Change billing rate" />
        <bt:String id="residCustomTabLabel" 
                    DefaultValue="Billable hours" />

        <bt:String id="residChangeToAssociateLabel" 
                    DefaultValue="Associate" />
      </bt:ShortStrings>

      <bt:LongStrings>
        <bt:String id="residVersionOverrideDesc" 
                    DefaultValue="Version override description" />

        <bt:String id="residChangeToAssociateDesc" 
                    DefaultValue="Change to the associate billing rate: $127/hr" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

## <a name="see-also"></a><span data-ttu-id="3191a-123">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3191a-123">See also</span></span>

- [<span data-ttu-id="3191a-124">Outlook 加载项清单</span><span class="sxs-lookup"><span data-stu-id="3191a-124">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="3191a-125">在清单中定义外接程序命令</span><span class="sxs-lookup"><span data-stu-id="3191a-125">Define add-in commands in your manifest</span></span>](../develop/define-add-in-commands.md)
- [<span data-ttu-id="3191a-126">Outlook 模块扩展计酬时间示例</span><span class="sxs-lookup"><span data-stu-id="3191a-126">Outlook module extensions Billable hours sample</span></span>](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ModuleExtension)

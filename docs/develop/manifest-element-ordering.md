---
title: 如何查找清单元素的正确顺序
description: 了解如何查找在父元素中放置子元素的正确顺序。
ms.date: 08/22/2019
localization_priority: Normal
ms.openlocfilehash: 6e88ee4f6eac09d4af8bc79a53f5faf890551a15
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111051"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>如何查找清单元素的正确顺序

Office 外接程序清单中的 XML 元素必须位于正确父元素下，*且*在父元素下以特定的相对顺序存在。

所需的排序在 [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 文件夹的 XSD 文件中指定。 XSD 文件分类存放在对应任务窗格、内容和邮件三类外接程序的子文件夹中。

例如，在 `<OfficeApp>` 元素中，`<Id>`、`<Version>`、`<ProviderName>` 必须按此顺序出现。 如果添加了 `<AlternateId>` 元素，则其必须位于 `<Id>` 和 `<Version>` 元素之间。 如果任何元素的顺序出错，清单将无效并且你的外接程序将无法加载。

> [!NOTE]
> [Office 外接程序指令清单中的验证](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest)器在元素的顺序不正确时使用相同的错误消息，与元素位于错误父项下时相同。 该错误消息会提示子元素不是父元素的有效子级。 如果出现此类错误，而子元素的参考文档却指示它对父级*是*有效的，则问题很可能是子级的放置顺序出现了错误。

以下各节按它们必须出现的顺序显示清单元素。 取决`type`于`<OfficeApp>`元素的属性是`TaskPaneApp`、 `ContentApp`还是，也`MailApp`有不同之处。 为了防止这些部分变得过于复杂，高度复杂`<VersionOverrides>`的元素将分解为单独的部分。

> [!Note]
> 并不是所有显示的元素都是必需的。 如果某个`minOccurs`元素的值在[架构](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)中为**0** ，则该元素是可选的。

## <a name="basic-task-pane-add-in-element-ordering"></a>基本任务窗格加载项元素排序

```xml
<OfficeApp xsi:type="TaskPaneApp">
    <Id>
    <AlternateID>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl>
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
        <Sets>
            <Set>
        <Methods>
            <Method>
    <DefaultSettings>
        <SourceLocation>
            <Override>
    <Permissions>
    <Dictionary>
        <TargetDialects>
        <QueryUri>
        <CitationText>
        <DictionaryName>
        <DictionaryHomePage>
    <VersionOverrides>*
```

\*有关 VersionOverrides 的子元素的排序，请参阅[VersionOverrides 内的任务窗格加载项元素排序](#task-pane-add-in-element-ordering-within-versionoverrides)。

## <a name="basic-mail-add-in-element-ordering"></a>基本邮件加载项元素排序

```xml
<OfficeApp xsi:type="MailApp">
    <Id>
    <AlternateId>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl>
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
    <Sets>
        <Set>
    <FormSettings>
        <Form>
        <DesktopSettings>
            <SourceLocation>
            <RequestedHeight>
        <TabletSettings>
            <SourceLocation>
            <RequestedHeight>
        <PhoneSettings>
            <SourceLocation>
    <Permissions>
    <Rule>
    <DisableEntityHighlighting>
    <VersionOverrides>*
```

\*有关 VersionOverrides 的子元素排序，请参阅[VersionOverrides. 1.0 中的邮件外接程序元素排序](#mail-add-in-element-ordering-within-versionoverrides-ver-10)和[VersionOverrides Ver 中的邮件加载项元素排序1.1。](#mail-add-in-element-ordering-within-versionoverrides-ver-11)

## <a name="basic-content-add-in-element-ordering"></a>基本内容加载项元素排序

```xml
<OfficeApp xsi:type="ContentApp">
    <Id>
    <AlternateId>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl >
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
    <Sets>
        <Set>
    <Methods>
        <Method>
    <DefaultSettings>
        <SourceLocation>
            <Override>
    <RequestedWidth>
    <RequestedHeight>
    <Permissions>
    <AllowSnapshot>
    <VersionOverrides>*
```

\*有关 VersionOverrides 的子元素的排序，请参阅[VersionOverrides 内的内容加载项元素排序](#content-add-in-element-ordering-within-versionoverrides)。

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a>VersionOverrides 中的任务窗格加载项元素排序

```xml
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
      <Hosts>
        <Host>
            <AllFormFactors>
            <ExtensionPoint>
                <Script>
                    <SourceLocation>
                <Page>
                    <SourceLocation>
                <Metadata>
                    <SourceLocation>
                <Namespace>
            <DesktopFormFactor>
            <GetStarted>
                <Title>
                <Description>
                <LearnMoreUrl>
            <FunctionFile>
            <ExtensionPoint>
                <OfficeTab>
                    <Group>
                        <Label>
                        <Icon>
                            <Image>
                        <Control>
                        <Label>
                        <Supertip>
                            <Title>
                            <Description>
                        <Icon>
                            <Image>  
                        <Action>
                            <TaskpaneId>
                            <SourceLocation>
                            <Title>
                            <FunctionName>
                        <Items>
                            <Item>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                <CustomTab>
                    <Group>
                        <Label>
                        <Icon>
                            <Image>
                        <Control>
                        <Label>
                        <Supertip>
                            <Title>
                            <Description>
                        <Icon>
                            <Image>  
                        <Action>
                            <TaskpaneId>
                            <SourceLocation>
                            <Title>
                            <FunctionName>
                        <Items>
                            <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                    <Label>
                <OfficeMenu>
                    <Control>
                        <Label>
                        <Supertip>
                            <Title>
                            <Description>
                        <Icon>
                            <Image>  
                        <Action>
                            <TaskpaneId>
                            <SourceLocation>
                            <Title>
                            <FunctionName>
                        <Items>
                            <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
        <Runtimes>
            <Runtime>
        <Resources>
            <Images>
                <Image>
                    <Override>
            <Urls>
                <Url>
                    <Override>
            <ShortStrings>
                <String>
                    <Override>
            <LongStrings>
                <String>
                    <Override>
        <WebApplicationInfo>
            <Id>
            <MsaId>
            <Resource>
            <Scopes>
                <Scope>
            <Authorizations>
                <Authorization>
                    <Resource>
                    <Scopes>
                        <Scope>
        <EquivalentAddins>
            <EquivalentAddin>
                <ProgId>
                <DisplayName>
                <FileName>
                <Type>
```

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a>VersionOverrides Ver 中的邮件加载项元素排序。 1.0

```xml
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
    <Hosts>
        <Host>
            <DesktopFormFactor>
            <ExtensionPoint>
                <OfficeTab>
                    <Group>
                        <Label>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>
                            <Action>
                                <SourceLocation>
                                <FunctionName>
                <CustomTab>
                    <Group>
                        <Label>
                        <Icon>
                            <Image>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
                    <Label>
                <OfficeMenu>
                    <Control>
                        <Label>
                        <Supertip>
                            <Title>
                            <Description>
                        <Icon>
                            <Image>
                        <Action>
                            <TaskpaneId>
                            <SourceLocation>
                            <Title>
                            <FunctionName>
                        <Items>
                            <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
    <Resources>
        <Images>
            <Image>
                <Override>
        <Urls>
            <Url>
                <Override>
        <ShortStrings>
            <String>
                <Override>
        <LongStrings>
            <String>
                <Override>
    <VersionOverrides>*
```

\*具有`type`值`VersionOverridesV1_1`（而不是`VersionOverridesV1_0`）的 VersionOverrides 可以嵌套在外部 VersionOverrides 的末尾。 有关中`VersionOverridesV1_1`的元素排序，请参阅[VersionOverrides 1.1 Ver 中的邮件加载项元素排序](#mail-add-in-element-ordering-within-versionoverrides-ver-11)。

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a>VersionOverrides Ver 中的邮件加载项元素排序。 1.1

```xml
<VersionOverrides>
    <Description>
    <Requirements>
    <Sets>
        <Set>
    <Hosts>
    <Host>
        <DesktopFormFactor>
        <ExtensionPoint>
            <OfficeTab>
                <Group>
                    <Label>
                    <Control>
                        <Label>
                        <Supertip>
                            <Title>
                            <Description>
                        <Icon>
                            <Image>
                        <Action>
                            <SourceLocation>
                            <FunctionName>
            <CustomTab>
                <Group>
                    <Label>
                    <Icon>
                        <Image>
                    <Control>
                        <Label>
                        <Supertip>
                            <Title>
                            <Description>
                        <Icon>
                            <Image>  
                        <Action>
                            <TaskpaneId>
                            <SourceLocation>
                            <Title>
                            <FunctionName>
                        <Items>
                            <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                <Label>
            <OfficeMenu>
                <Control>
                    <Label>
                    <Supertip>
                        <Title>
                        <Description>
                    <Icon>
                        <Image>  
                    <Action>
                        <TaskpaneId>
                        <SourceLocation>
                        <Title>
                        <FunctionName>
                    <Items>
                        <Item>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                                <SourceLocation>
            <SourceLocation>
            <Label>
            <CommandSurface>
    <Resources>
        <Images>
            <Image>
                <Override>
        <Urls>
            <Url>
                <Override>
        <ShortStrings>
            <String>
                <Override>
        <LongStrings>
            <String>
                <Override>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="content-add-in-element-ordering-within-versionoverrides"></a>VersionOverrides 中的内容加载项元素排序

```xml
<VersionOverrides>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="see-also"></a>另请参阅

- [Office 外接程序清单的架构参考 (v1.1)](../develop/add-in-manifests.md)

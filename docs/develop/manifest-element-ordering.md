---
title: 如何查找清单元素的正确顺序
description: 了解如何查找在父元素中放置子元素的正确顺序。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: d418f796592a0e4c247e717a5ce75d1c40c18d79
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302572"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a><span data-ttu-id="87b56-103">如何查找清单元素的正确顺序</span><span class="sxs-lookup"><span data-stu-id="87b56-103">How to find the proper order of manifest elements</span></span>

<span data-ttu-id="87b56-104">Office 外接程序清单中的 XML 元素必须位于正确父元素下，*且*在父元素下以特定的相对顺序存在。</span><span class="sxs-lookup"><span data-stu-id="87b56-104">The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.</span></span>

<span data-ttu-id="87b56-105">所需的排序在 [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 文件夹的 XSD 文件中指定。</span><span class="sxs-lookup"><span data-stu-id="87b56-105">The required ordering is specified in the XSD files in the [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) folder.</span></span> <span data-ttu-id="87b56-106">XSD 文件分类存放在对应任务窗格、内容和邮件三类外接程序的子文件夹中。</span><span class="sxs-lookup"><span data-stu-id="87b56-106">The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.</span></span>

<span data-ttu-id="87b56-107">例如，在 `<OfficeApp>` 元素中，`<Id>`、`<Version>`、`<ProviderName>` 必须按此顺序出现。</span><span class="sxs-lookup"><span data-stu-id="87b56-107">For example, in the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order.</span></span> <span data-ttu-id="87b56-108">如果添加了 `<AlternateId>` 元素，则其必须位于 `<Id>` 和 `<Version>` 元素之间。</span><span class="sxs-lookup"><span data-stu-id="87b56-108">If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element.</span></span> <span data-ttu-id="87b56-109">如果任何元素的顺序出错，清单将无效并且你的外接程序将无法加载。</span><span class="sxs-lookup"><span data-stu-id="87b56-109">Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.</span></span>

> [!NOTE]
> <span data-ttu-id="87b56-110">[Office-工具箱中的验证](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-toolbox)程序在元素的顺序不正确时使用相同的错误消息, 如同元素位于错误父项下时的情况一样。</span><span class="sxs-lookup"><span data-stu-id="87b56-110">The [validator within office-toolbox](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-toolbox) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent.</span></span> <span data-ttu-id="87b56-111">该错误消息会提示子元素不是父元素的有效子级。</span><span class="sxs-lookup"><span data-stu-id="87b56-111">The error says the child element is not a valid child of the parent element.</span></span> <span data-ttu-id="87b56-112">如果出现此类错误，而子元素的参考文档却指示它对父级*是*有效的，则问题很可能是子级的放置顺序出现了错误。</span><span class="sxs-lookup"><span data-stu-id="87b56-112">If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.</span></span>

<span data-ttu-id="87b56-113">以下各节按它们必须出现的顺序显示清单元素。</span><span class="sxs-lookup"><span data-stu-id="87b56-113">The following sections show the manifest elements in the order in which they must appear.</span></span> <span data-ttu-id="87b56-114">取决`type`于`<OfficeApp>`元素的属性是`TaskPaneApp`、 `ContentApp`还是, 也`MailApp`会稍有不同。</span><span class="sxs-lookup"><span data-stu-id="87b56-114">There are slight differences depending on whether the `type` attribute of the `<OfficeApp>` element is `TaskPaneApp`, `ContentApp`, or `MailApp`.</span></span> <span data-ttu-id="87b56-115">为了防止这些部分变得过于复杂, 高度复杂`<VersionOverrides>`的元素将分解为单独的部分。</span><span class="sxs-lookup"><span data-stu-id="87b56-115">To keep these sections from becoming too unwieldy, the highly complex `<VersionOverrides>` element is broken out into separate sections.</span></span>

> [!Note]
> <span data-ttu-id="87b56-116">并非所有元素都显示为强制性的。</span><span class="sxs-lookup"><span data-stu-id="87b56-116">Not all of the elements show are mandatory.</span></span> <span data-ttu-id="87b56-117">如果某个`minOccurs`元素的值在[架构](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)中为**0** , 则该元素是可选的。</span><span class="sxs-lookup"><span data-stu-id="87b56-117">If the `minOccurs` value for a element is **0** in the [schema](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), the element is optional.</span></span>

## <a name="basic-task-pane-add-in-element-ordering"></a><span data-ttu-id="87b56-118">基本任务窗格加载项元素排序</span><span class="sxs-lookup"><span data-stu-id="87b56-118">Basic task pane add-in element ordering</span></span>

```
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

<span data-ttu-id="87b56-119">\*有关 VersionOverrides 的子元素的排序, 请参阅[VersionOverrides 内的任务窗格加载项元素排序](#task-pane-add-in-element-ordering-within-versionoverrides)。</span><span class="sxs-lookup"><span data-stu-id="87b56-119">\*See [Task pane add-in element ordering within VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-mail-add-in-element-ordering"></a><span data-ttu-id="87b56-120">基本邮件加载项元素排序</span><span class="sxs-lookup"><span data-stu-id="87b56-120">Basic mail add-in element ordering</span></span>

```
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

<span data-ttu-id="87b56-121">\*有关 VersionOverrides 的子元素排序, 请参阅[VersionOverrides. 1.0 中的邮件外接程序元素排序](#mail-add-in-element-ordering-within-versionoverrides-ver-10)和[VersionOverrides Ver 中的邮件加载项元素排序1.1。](#mail-add-in-element-ordering-within-versionoverrides-ver-11)</span><span class="sxs-lookup"><span data-stu-id="87b56-121">\*See [Mail add-in element ordering within VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) and [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-content-add-in-element-ordering"></a><span data-ttu-id="87b56-122">基本内容加载项元素排序</span><span class="sxs-lookup"><span data-stu-id="87b56-122">Basic content add-in element ordering</span></span>

```
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
    <VersionOverrides>
```

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="87b56-123">VersionOverrides 中的任务窗格加载项元素排序</span><span class="sxs-lookup"><span data-stu-id="87b56-123">Task pane add-in element ordering within VersionOverrides</span></span>

```
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

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a><span data-ttu-id="87b56-124">VersionOverrides Ver 中的邮件加载项元素排序。</span><span class="sxs-lookup"><span data-stu-id="87b56-124">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="87b56-125">1.0</span><span class="sxs-lookup"><span data-stu-id="87b56-125">1.0</span></span>

```
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

<span data-ttu-id="87b56-126">\*具有`type`值`VersionOverridesV1_1`(而不是`VersionOverridesV1_0`) 的 VersionOverrides 可以嵌套在外部 VersionOverrides 的末尾。</span><span class="sxs-lookup"><span data-stu-id="87b56-126">\* A VersionOverrides with `type` value `VersionOverridesV1_1`, instead of `VersionOverridesV1_0`, can be nested at the end of the outer VersionOverrides.</span></span> <span data-ttu-id="87b56-127">有关中`VersionOverridesV1_1`的元素排序, 请参阅[VersionOverrides 1.1 Ver 中的邮件加载项元素排序](#mail-add-in-element-ordering-within-versionoverrides-ver-11)。</span><span class="sxs-lookup"><span data-stu-id="87b56-127">See [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of elements in `VersionOverridesV1_1`.</span></span>

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a><span data-ttu-id="87b56-128">VersionOverrides Ver 中的邮件加载项元素排序。</span><span class="sxs-lookup"><span data-stu-id="87b56-128">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="87b56-129">1.1</span><span class="sxs-lookup"><span data-stu-id="87b56-129">1.1</span></span>

```
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

## <a name="see-also"></a><span data-ttu-id="87b56-130">另请参阅</span><span class="sxs-lookup"><span data-stu-id="87b56-130">See also</span></span>

- [<span data-ttu-id="87b56-131">Office 外接程序清单的架构参考 (v1.1)</span><span class="sxs-lookup"><span data-stu-id="87b56-131">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)

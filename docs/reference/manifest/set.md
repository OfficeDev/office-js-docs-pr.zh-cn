# <a name="set-element"></a><span data-ttu-id="53371-101">Set 元素</span><span class="sxs-lookup"><span data-stu-id="53371-101">Set element</span></span>

<span data-ttu-id="53371-102">指定来自适用于 Office 的 JavaScript API 的要求集合，Office 外接程序需要该集才能激活。</span><span class="sxs-lookup"><span data-stu-id="53371-102">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="53371-103">**加载项类型：** Content、Task pane、mail</span><span class="sxs-lookup"><span data-stu-id="53371-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="53371-104">句法</span><span class="sxs-lookup"><span data-stu-id="53371-104">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="53371-105">包含在</span><span class="sxs-lookup"><span data-stu-id="53371-105">Contained in:</span></span>

[<span data-ttu-id="53371-106">集</span><span class="sxs-lookup"><span data-stu-id="53371-106">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="53371-107">属性</span><span class="sxs-lookup"><span data-stu-id="53371-107">Attributes</span></span>

|<span data-ttu-id="53371-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="53371-108">**Attribute**</span></span>|<span data-ttu-id="53371-109">**类型**</span><span class="sxs-lookup"><span data-stu-id="53371-109">**Type**</span></span>|<span data-ttu-id="53371-110">**必需**</span><span class="sxs-lookup"><span data-stu-id="53371-110">**Required**</span></span>|<span data-ttu-id="53371-111">**说明**</span><span class="sxs-lookup"><span data-stu-id="53371-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="53371-112">名称</span><span class="sxs-lookup"><span data-stu-id="53371-112">Name</span></span>|<span data-ttu-id="53371-113">String</span><span class="sxs-lookup"><span data-stu-id="53371-113">string</span></span>|<span data-ttu-id="53371-114">必需</span><span class="sxs-lookup"><span data-stu-id="53371-114">required</span></span>|<span data-ttu-id="53371-115">[要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)名称。</span><span class="sxs-lookup"><span data-stu-id="53371-115">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="53371-116">MinVersion</span><span class="sxs-lookup"><span data-stu-id="53371-116">MinVersion</span></span>|<span data-ttu-id="53371-117">String</span><span class="sxs-lookup"><span data-stu-id="53371-117">string</span></span>|<span data-ttu-id="53371-118">可选</span><span class="sxs-lookup"><span data-stu-id="53371-118">optional</span></span>|<span data-ttu-id="53371-p101">指定加载项所需的 API 集的最低版本。如果 **DefaultMinVersion** 的值已在父 [Sets](sets.md) 元素中指定，则替代该值。</span><span class="sxs-lookup"><span data-stu-id="53371-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="53371-121">备注</span><span class="sxs-lookup"><span data-stu-id="53371-121">Remarks</span></span>

<span data-ttu-id="53371-122">欲知要求集的详细信息，请参阅 [Office 版本和要求集合](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="53371-122">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="53371-123">欲知 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置要求元素](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="53371-123">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="53371-124">对于邮件加载项，只有一个 `"Mailbox"` 要求集合可用。</span><span class="sxs-lookup"><span data-stu-id="53371-124">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="53371-125">此要求集合包含整个 outlook 邮件加载项中支持的 API 子集合，您必须指定邮件加载项清单中设置的 `"Mailbox"` 要求(对于内容和任务窗格加载项，它不是可选的)。</span><span class="sxs-lookup"><span data-stu-id="53371-125">Important  For mail add-ins, there is only one   requirement set available. This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins). Also, you can't declare support for specific methods in mail add-ins.</span></span> <span data-ttu-id="53371-126">此外，无法声明支持邮件加载项中的特定模式。</span><span class="sxs-lookup"><span data-stu-id="53371-126">Also, you can't declare support for specific methods in mail add-ins.</span></span>

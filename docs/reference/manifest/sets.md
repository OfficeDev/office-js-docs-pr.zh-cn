# <a name="sets-element"></a><span data-ttu-id="bdcd6-101">Sets 元素</span><span class="sxs-lookup"><span data-stu-id="bdcd6-101">Sets element</span></span>

<span data-ttu-id="bdcd6-102">指定适用于 Office 的 JavaScript API 的最小子集，Office 外接程序需要该子集才能激活。</span><span class="sxs-lookup"><span data-stu-id="bdcd6-102">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="bdcd6-103">**外接程序类型：** Content、Task pane、Mail</span><span class="sxs-lookup"><span data-stu-id="bdcd6-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bdcd6-104">语法</span><span class="sxs-lookup"><span data-stu-id="bdcd6-104">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="bdcd6-105">包含在</span><span class="sxs-lookup"><span data-stu-id="bdcd6-105">Contained in:</span></span>

[<span data-ttu-id="bdcd6-106">要求</span><span class="sxs-lookup"><span data-stu-id="bdcd6-106">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="bdcd6-107">可以包含</span><span class="sxs-lookup"><span data-stu-id="bdcd6-107">Can contain:</span></span>

[<span data-ttu-id="bdcd6-108">Set</span><span class="sxs-lookup"><span data-stu-id="bdcd6-108">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="bdcd6-109">属性</span><span class="sxs-lookup"><span data-stu-id="bdcd6-109">Attributes</span></span>

|<span data-ttu-id="bdcd6-110">**属性**</span><span class="sxs-lookup"><span data-stu-id="bdcd6-110">**Attribute**</span></span>|<span data-ttu-id="bdcd6-111">**类型**</span><span class="sxs-lookup"><span data-stu-id="bdcd6-111">**Type**</span></span>|<span data-ttu-id="bdcd6-112">**是否必需**</span><span class="sxs-lookup"><span data-stu-id="bdcd6-112">**Required**</span></span>|<span data-ttu-id="bdcd6-113">**说明**</span><span class="sxs-lookup"><span data-stu-id="bdcd6-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bdcd6-114">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="bdcd6-114">DefaultMinVersion</span></span>|<span data-ttu-id="bdcd6-115">字符串</span><span class="sxs-lookup"><span data-stu-id="bdcd6-115">string</span></span>|<span data-ttu-id="bdcd6-116">可选</span><span class="sxs-lookup"><span data-stu-id="bdcd6-116">optional</span></span>|<span data-ttu-id="bdcd6-p101">为所有子 **Set** 元素指定默认的 [MinVersion](set.md) 属性值。默认值为“1.1”。</span><span class="sxs-lookup"><span data-stu-id="bdcd6-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="bdcd6-119">备注</span><span class="sxs-lookup"><span data-stu-id="bdcd6-119">Remarks</span></span>

<span data-ttu-id="bdcd6-120">有关要求集的详细信息，请参阅 [Office 版本和要求集合](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="bdcd6-120">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="bdcd6-121">有关 **集合** 元素的 **MinVersion** 属性和 **集合** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置要求元素](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="bdcd6-121">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>


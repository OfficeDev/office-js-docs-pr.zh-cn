# <a name="requirements-element"></a><span data-ttu-id="2427e-101">要求元素</span><span class="sxs-lookup"><span data-stu-id="2427e-101">Requirements element</span></span>

<span data-ttu-id="2427e-102">指定适用于 Office 的 JavaScript API 要求（[要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)和/或 方法）的最小集，Office 外接程序需要该集才能激活。</span><span class="sxs-lookup"><span data-stu-id="2427e-102">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="2427e-103">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="2427e-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2427e-104">语法</span><span class="sxs-lookup"><span data-stu-id="2427e-104">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="2427e-105">包含在</span><span class="sxs-lookup"><span data-stu-id="2427e-105">Contained in:</span></span>

[<span data-ttu-id="2427e-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="2427e-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="2427e-107">可以包含</span><span class="sxs-lookup"><span data-stu-id="2427e-107">Can contain:</span></span>

|<span data-ttu-id="2427e-108">**元素**</span><span class="sxs-lookup"><span data-stu-id="2427e-108">**Element**</span></span>|<span data-ttu-id="2427e-109">**内容**</span><span class="sxs-lookup"><span data-stu-id="2427e-109">**Content**</span></span>|<span data-ttu-id="2427e-110">**邮件**</span><span class="sxs-lookup"><span data-stu-id="2427e-110">**Mail**</span></span>|<span data-ttu-id="2427e-111">**任务窗格**</span><span class="sxs-lookup"><span data-stu-id="2427e-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="2427e-112">集</span><span class="sxs-lookup"><span data-stu-id="2427e-112">Sets</span></span>](sets.md)|<span data-ttu-id="2427e-113">x</span><span class="sxs-lookup"><span data-stu-id="2427e-113">x</span></span>|<span data-ttu-id="2427e-114">x</span><span class="sxs-lookup"><span data-stu-id="2427e-114">x</span></span>|<span data-ttu-id="2427e-115">x</span><span class="sxs-lookup"><span data-stu-id="2427e-115">x</span></span>|
|[<span data-ttu-id="2427e-116">方法</span><span class="sxs-lookup"><span data-stu-id="2427e-116">Methods</span></span>](methods.md)|<span data-ttu-id="2427e-117">x</span><span class="sxs-lookup"><span data-stu-id="2427e-117">x</span></span>||<span data-ttu-id="2427e-118">x</span><span class="sxs-lookup"><span data-stu-id="2427e-118">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="2427e-119">备注</span><span class="sxs-lookup"><span data-stu-id="2427e-119">Remarks</span></span>

<span data-ttu-id="2427e-120">有关要求集的详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="2427e-120">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>


# <a name="defaultsettings-element"></a><span data-ttu-id="98315-101">DefaultSettings 元素</span><span class="sxs-lookup"><span data-stu-id="98315-101">DefaultSettings element</span></span>

<span data-ttu-id="98315-102">指定内容或任务窗格外接程序的默认源位置和其他默认设置。</span><span class="sxs-lookup"><span data-stu-id="98315-102">Specifies the default source location and other default settings for your content or task pane add-in .</span></span>

<span data-ttu-id="98315-103">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="98315-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="98315-104">语法</span><span class="sxs-lookup"><span data-stu-id="98315-104">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="98315-105">包含在</span><span class="sxs-lookup"><span data-stu-id="98315-105">Contained in:</span></span>

[<span data-ttu-id="98315-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="98315-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="98315-107">可以包含</span><span class="sxs-lookup"><span data-stu-id="98315-107">Can contain:</span></span>

|<span data-ttu-id="98315-108">**元素**</span><span class="sxs-lookup"><span data-stu-id="98315-108">**Element**</span></span>|<span data-ttu-id="98315-109">**内容**</span><span class="sxs-lookup"><span data-stu-id="98315-109">**Content**</span></span>|<span data-ttu-id="98315-110">**邮件**</span><span class="sxs-lookup"><span data-stu-id="98315-110">**Mail**</span></span>|<span data-ttu-id="98315-111">**任务窗格**</span><span class="sxs-lookup"><span data-stu-id="98315-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="98315-112">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="98315-112">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="98315-113">x</span><span class="sxs-lookup"><span data-stu-id="98315-113">x</span></span>||<span data-ttu-id="98315-114">x</span><span class="sxs-lookup"><span data-stu-id="98315-114">x</span></span>|
|[<span data-ttu-id="98315-115">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="98315-115">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="98315-116">x</span><span class="sxs-lookup"><span data-stu-id="98315-116">x</span></span>|||
|[<span data-ttu-id="98315-117">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="98315-117">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="98315-118">x</span><span class="sxs-lookup"><span data-stu-id="98315-118">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="98315-119">备注</span><span class="sxs-lookup"><span data-stu-id="98315-119">Remarks</span></span>

<span data-ttu-id="98315-120">**DefaultSettings** 元素中的源位置和其他设置仅应用于内容和任务窗格外接程序。对于邮件外接程序，在 [FormSettings](formsettings.md) 元素中指定源文件的默认位置和其他默认设置。</span><span class="sxs-lookup"><span data-stu-id="98315-120">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>


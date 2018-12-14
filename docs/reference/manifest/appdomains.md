# <a name="appdomains-element"></a><span data-ttu-id="4c5a0-101">AppDomains 元素</span><span class="sxs-lookup"><span data-stu-id="4c5a0-101">AppDomains element</span></span>

<span data-ttu-id="4c5a0-p101">列出了除 Office 外接程序用于加载页面的 SourceLocation 元素中指定的域之外的所有域。对于每个其他域，指定 AppDomain 元素。</span><span class="sxs-lookup"><span data-stu-id="4c5a0-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="4c5a0-104">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="4c5a0-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4c5a0-105">语法</span><span class="sxs-lookup"><span data-stu-id="4c5a0-105">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="4c5a0-106">每个 **AppDomain** 元素的值都必须包括协议（如 `<AppDomain>https://myappdomain<AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="4c5a0-106">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="4c5a0-107">包含于</span><span class="sxs-lookup"><span data-stu-id="4c5a0-107">Contained in</span></span>

[<span data-ttu-id="4c5a0-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="4c5a0-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="4c5a0-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="4c5a0-109">Can contain</span></span>

[<span data-ttu-id="4c5a0-110">AppDomain</span><span class="sxs-lookup"><span data-stu-id="4c5a0-110">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="4c5a0-111">注释</span><span class="sxs-lookup"><span data-stu-id="4c5a0-111">Remarks</span></span>

<span data-ttu-id="4c5a0-112">默认情况下，外接程序可以加载与 [SourceLocation](sourcelocation.md) 元素中指定的位置位于同一个域中的任何页面。</span><span class="sxs-lookup"><span data-stu-id="4c5a0-112">By default, your add-in can load any page that is in the same domain as the location specified in the SourceLocation element. To load pages that are not in the same domain as the add-in, specify the domains by using the AppDomains and AppDomain elements. This element can't be empty.</span></span> <span data-ttu-id="4c5a0-113">要加载与外接程序位于不同域中的页面，可以使用 **AppDomains** 和 **AppDomain** 元素来指定域。</span><span class="sxs-lookup"><span data-stu-id="4c5a0-113">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="4c5a0-114">此元素不能为空。</span><span class="sxs-lookup"><span data-stu-id="4c5a0-114">This element can't be empty.</span></span>

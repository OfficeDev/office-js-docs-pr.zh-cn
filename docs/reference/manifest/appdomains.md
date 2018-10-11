# <a name="appdomains-element"></a><span data-ttu-id="ff3e3-101">AppDomains 元素</span><span class="sxs-lookup"><span data-stu-id="ff3e3-101">AppDomains element</span></span>

<span data-ttu-id="ff3e3-p101">列出了除 Office 加载项用于加载页面的 SourceLocation 元素中指定的域以外的所有域。对于每个其他域，指定 AppDomain 元素。</span><span class="sxs-lookup"><span data-stu-id="ff3e3-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="ff3e3-104">**加载项类型：** Content、Task pane、Mail</span><span class="sxs-lookup"><span data-stu-id="ff3e3-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ff3e3-105">语法</span><span class="sxs-lookup"><span data-stu-id="ff3e3-105">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a><span data-ttu-id="ff3e3-106">包含在</span><span class="sxs-lookup"><span data-stu-id="ff3e3-106">Contained in:</span></span>

[<span data-ttu-id="ff3e3-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="ff3e3-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="ff3e3-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="ff3e3-108">Can contain:</span></span>

[<span data-ttu-id="ff3e3-109">AppDomain</span><span class="sxs-lookup"><span data-stu-id="ff3e3-109">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="ff3e3-110">备注</span><span class="sxs-lookup"><span data-stu-id="ff3e3-110">Remarks</span></span>

<span data-ttu-id="ff3e3-p102">默认情况下，加载项可以加载与 **SourceLocation** 元素中指定的位置位于同一个域中的任何页面。要加载不与加载项位于同一个域中的页面，请使用 **AppDomains** 和 **AppDomain** 元素来指定域。此元素不能为空。</span><span class="sxs-lookup"><span data-stu-id="ff3e3-p102">By default, your add-in can load any page that is in the same domain as the location specified in the **SourceLocation** element. To load pages that are not in the same domain as the add-in, specify the domains by using the **AppDomains** and **AppDomain** elements. This element can't be empty.</span></span> 

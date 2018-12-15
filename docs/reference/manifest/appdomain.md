# <a name="appdomain-element"></a><span data-ttu-id="64983-101">AppDomain 元素</span><span class="sxs-lookup"><span data-stu-id="64983-101">AppDomain element</span></span>

<span data-ttu-id="64983-102">指定将用于在外接程序窗口中加载页面的其他域。</span><span class="sxs-lookup"><span data-stu-id="64983-102">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="64983-103">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="64983-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="64983-104">语法</span><span class="sxs-lookup"><span data-stu-id="64983-104">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> <span data-ttu-id="64983-105">**AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain<AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="64983-105">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="64983-106">包含于</span><span class="sxs-lookup"><span data-stu-id="64983-106">Contained in</span></span>

[<span data-ttu-id="64983-107">AppDomains</span><span class="sxs-lookup"><span data-stu-id="64983-107">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="64983-108">注释</span><span class="sxs-lookup"><span data-stu-id="64983-108">Remarks</span></span>

<span data-ttu-id="64983-109">**AppDomain** 元素用于指定除在 [SourceLocation](sourcelocation.md) 元素中指定的域之外的任何其他域。</span><span class="sxs-lookup"><span data-stu-id="64983-109">The  AppDomains and **AppDomain** elements are used to specify any additional domains other than the one specified in the [SourceLocation element. For more information, see Office Add-ins XML manifest](sourcelocation.md).</span></span> <span data-ttu-id="64983-110">有关详细信息，请参阅 [Office 加载项 XML 清单](/office/dev/add-ins/develop/add-in-manifests)。</span><span class="sxs-lookup"><span data-stu-id="64983-110">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>

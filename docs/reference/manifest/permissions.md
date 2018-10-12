# <a name="permissions-element"></a><span data-ttu-id="a20aa-101">Permissions 元素</span><span class="sxs-lookup"><span data-stu-id="a20aa-101">Permissions element</span></span>

<span data-ttu-id="a20aa-102">指定 Office 加载项的 API 访问级别；应基于最低权限原则请求权限。</span><span class="sxs-lookup"><span data-stu-id="a20aa-102">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="a20aa-103">**加载项类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="a20aa-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a20aa-104">语法</span><span class="sxs-lookup"><span data-stu-id="a20aa-104">Syntax</span></span>

<span data-ttu-id="a20aa-105">对于内容和任务窗格加载项：</span><span class="sxs-lookup"><span data-stu-id="a20aa-105">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="a20aa-106">对于邮件加载项：</span><span class="sxs-lookup"><span data-stu-id="a20aa-106">For mail add-ins:</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="a20aa-107">包含在</span><span class="sxs-lookup"><span data-stu-id="a20aa-107">Contained in:</span></span>

[<span data-ttu-id="a20aa-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="a20aa-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="a20aa-109">说明</span><span class="sxs-lookup"><span data-stu-id="a20aa-109">Remarks</span></span>

<span data-ttu-id="a20aa-110">有关详细信息，请参阅[在内容和任务窗格加载项中请求 API 的使用权限](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)和[了解 Outlook 加载项权限](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="a20aa-110">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

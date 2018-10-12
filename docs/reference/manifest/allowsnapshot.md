# <a name="allowsnapshot-element"></a><span data-ttu-id="eeafa-101">AllowSnapshot 元素</span><span class="sxs-lookup"><span data-stu-id="eeafa-101">AllowSnapshot element</span></span>

<span data-ttu-id="eeafa-102">指定是否将内容外接程序的快照图像与主机文档一起保存。</span><span class="sxs-lookup"><span data-stu-id="eeafa-102">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="eeafa-103">**外接程序类型：** 内容</span><span class="sxs-lookup"><span data-stu-id="eeafa-103">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="eeafa-104">语法</span><span class="sxs-lookup"><span data-stu-id="eeafa-104">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="eeafa-105">包含在</span><span class="sxs-lookup"><span data-stu-id="eeafa-105">Contained in:</span></span>

[<span data-ttu-id="eeafa-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="eeafa-106">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="eeafa-107">说明</span><span class="sxs-lookup"><span data-stu-id="eeafa-107">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="eeafa-108">**AllowSnapshot** 默认是 `true`。</span><span class="sxs-lookup"><span data-stu-id="eeafa-108">**Security Note:AllowSnapshot** is true`true` by default.</span></span> <span data-ttu-id="eeafa-109">这样，用户在不支持 Office 外接程序的主机应用程序版本中打开文档时，即可看到该外接程序的图像，或者如果主机应用程序无法连接到托管外接程序的服务器时，会提供该外接程序的静态图像。</span><span class="sxs-lookup"><span data-stu-id="eeafa-109">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="eeafa-110">但是，这也意味着可以直接从托管该外接程序的文档访问显示在外接程序中的潜在敏感信息。</span><span class="sxs-lookup"><span data-stu-id="eeafa-110">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>


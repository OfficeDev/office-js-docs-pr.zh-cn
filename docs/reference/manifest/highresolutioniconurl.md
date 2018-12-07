# <a name="highresolutioniconurl-element"></a><span data-ttu-id="6c75e-101">HighResolutionIconUrl 元素</span><span class="sxs-lookup"><span data-stu-id="6c75e-101">HighResolutionIconUrl element</span></span>

<span data-ttu-id="6c75e-102">指定用于表示插入 UX 中的 Office 外接程序和高 DPI 屏幕上的 Office 应用商店的图像的 URL。</span><span class="sxs-lookup"><span data-stu-id="6c75e-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="6c75e-103">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="6c75e-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6c75e-104">语法</span><span class="sxs-lookup"><span data-stu-id="6c75e-104">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="6c75e-105">可以包含</span><span class="sxs-lookup"><span data-stu-id="6c75e-105">Can contain</span></span>

[<span data-ttu-id="6c75e-106">替代</span><span class="sxs-lookup"><span data-stu-id="6c75e-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="6c75e-107">属性</span><span class="sxs-lookup"><span data-stu-id="6c75e-107">Attributes</span></span>

|<span data-ttu-id="6c75e-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="6c75e-108">**Attribute**</span></span>|<span data-ttu-id="6c75e-109">**类型**</span><span class="sxs-lookup"><span data-stu-id="6c75e-109">**Type**</span></span>|<span data-ttu-id="6c75e-110">**必需**</span><span class="sxs-lookup"><span data-stu-id="6c75e-110">**Required**</span></span>|<span data-ttu-id="6c75e-111">**说明**</span><span class="sxs-lookup"><span data-stu-id="6c75e-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6c75e-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="6c75e-112">DefaultValue</span></span>|<span data-ttu-id="6c75e-113">字符串 (URL)</span><span class="sxs-lookup"><span data-stu-id="6c75e-113">string (URL)</span></span>|<span data-ttu-id="6c75e-114">必需</span><span class="sxs-lookup"><span data-stu-id="6c75e-114">required</span></span>|<span data-ttu-id="6c75e-115">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="6c75e-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="6c75e-116">注解</span><span class="sxs-lookup"><span data-stu-id="6c75e-116">Remarks</span></span>

<span data-ttu-id="6c75e-p101">对于邮件外接程序，图标显示在“**文件**” > “**管理外接程序**”UI 中。对于内容或任务窗格外接程序，图标显示在“**插入**” > “**外接程序**”UI 中。</span><span class="sxs-lookup"><span data-stu-id="6c75e-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="6c75e-119">图像必须采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。</span><span class="sxs-lookup"><span data-stu-id="6c75e-119">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="6c75e-120">对于内容和任务窗格应用程序，推荐的图像分辨率是 64 x 64 像素。</span><span class="sxs-lookup"><span data-stu-id="6c75e-120">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="6c75e-121">对于邮件应用程序，图像必须是 128 x 128 像素。</span><span class="sxs-lookup"><span data-stu-id="6c75e-121">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="6c75e-122">有关详细信息，请参阅[在 AppSource 和 Office 中创建有效的应用一览](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)中的_为你的应用创建一致的视觉标识_部分。</span><span class="sxs-lookup"><span data-stu-id="6c75e-122">For more information, see the section  Create a consistent visual identity for your app in Create effective Office Store apps and add-ins.</span></span>

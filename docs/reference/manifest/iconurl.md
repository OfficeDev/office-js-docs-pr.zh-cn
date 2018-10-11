# <a name="iconurl-element"></a><span data-ttu-id="21dac-101">IconUrl 元素</span><span class="sxs-lookup"><span data-stu-id="21dac-101">IconUrl element</span></span>

<span data-ttu-id="21dac-102">指定用于表示插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。</span><span class="sxs-lookup"><span data-stu-id="21dac-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="21dac-103">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="21dac-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="21dac-104">句法</span><span class="sxs-lookup"><span data-stu-id="21dac-104">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="21dac-105">可以包含</span><span class="sxs-lookup"><span data-stu-id="21dac-105">Can contain:</span></span>

[<span data-ttu-id="21dac-106">覆盖</span><span class="sxs-lookup"><span data-stu-id="21dac-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="21dac-107">属性</span><span class="sxs-lookup"><span data-stu-id="21dac-107">Attributes</span></span>

|<span data-ttu-id="21dac-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="21dac-108">**Attribute**</span></span>|<span data-ttu-id="21dac-109">**类型**</span><span class="sxs-lookup"><span data-stu-id="21dac-109">**Type**</span></span>|<span data-ttu-id="21dac-110">**必需**</span><span class="sxs-lookup"><span data-stu-id="21dac-110">**Required**</span></span>|<span data-ttu-id="21dac-111">**说明**</span><span class="sxs-lookup"><span data-stu-id="21dac-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="21dac-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="21dac-112">DefaultValue</span></span>|<span data-ttu-id="21dac-113">string</span><span class="sxs-lookup"><span data-stu-id="21dac-113">string</span></span>|<span data-ttu-id="21dac-114">必需</span><span class="sxs-lookup"><span data-stu-id="21dac-114">required</span></span>|<span data-ttu-id="21dac-115">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="21dac-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="21dac-116">备注</span><span class="sxs-lookup"><span data-stu-id="21dac-116">Remarks</span></span>

<span data-ttu-id="21dac-p101">对于邮件外接程序，该图标显示在**文件** > **管理外接程序**UI (Outlook) 或**设置** > **管理外接程序**UI (Outlook Web App) 中。对于内容或任务窗格外接程序，图标显示在**插入** > **外接程序**UI 中。对于所有外接程序类型，如果您将外接程序发布到 Office 应用商店，则该图标也将用于 Office 应用商店网站上。</span><span class="sxs-lookup"><span data-stu-id="21dac-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook Web App). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span></span>

<span data-ttu-id="21dac-120">图像必须是下列文件格式之一： GIF、 JPG、 PNG、 EXIF、 BMP 或 TIFF。</span><span class="sxs-lookup"><span data-stu-id="21dac-120">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="21dac-121">用于内容和任务窗格应用程序的指定图像，必须是 32 x 32 像素。</span><span class="sxs-lookup"><span data-stu-id="21dac-121">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="21dac-122">用于邮件应用程序的图像，必须是 64 x 64 像素。</span><span class="sxs-lookup"><span data-stu-id="21dac-122">For mail apps, the image must be 64 x 64 pixels.</span></span> <span data-ttu-id="21dac-123">您还应使用 [  HighResolutionIconUrl ](highresolutioniconurl.md) 元素，指定用于在高 DPI 屏幕上运行的 Office 主机应用程序的图标</span><span class="sxs-lookup"><span data-stu-id="21dac-123">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="21dac-124">有关详细信息，请参阅 [创建有效列表在 AppSource 中，在 Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity) 中 _创建一致的视觉标识您的应用程序_一节。</span><span class="sxs-lookup"><span data-stu-id="21dac-124">For more information, see the section  Create a consistent visual identity for your app in Create effective Office Store apps and add-ins.</span></span>

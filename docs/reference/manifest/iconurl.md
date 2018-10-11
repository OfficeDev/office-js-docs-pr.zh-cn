# <a name="iconurl-element"></a>IconUrl 元素

指定用于表示插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。

**外接程序类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>句法

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>可以包含

[覆盖](override.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string|必需|指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|

## <a name="remarks"></a>备注

对于邮件外接程序，该图标显示在**文件** > **管理外接程序**UI (Outlook) 或**设置** > **管理外接程序**UI (Outlook Web App) 中。对于内容或任务窗格外接程序，图标显示在**插入** > **外接程序**UI 中。对于所有外接程序类型，如果您将外接程序发布到 Office 应用商店，则该图标也将用于 Office 应用商店网站上。

图像必须是下列文件格式之一： GIF、 JPG、 PNG、 EXIF、 BMP 或 TIFF。 用于内容和任务窗格应用程序的指定图像，必须是 32 x 32 像素。 用于邮件应用程序的图像，必须是 64 x 64 像素。 您还应使用 [  HighResolutionIconUrl ](highresolutioniconurl.md) 元素，指定用于在高 DPI 屏幕上运行的 Office 主机应用程序的图标 有关详细信息，请参阅 [创建有效列表在 AppSource 中，在 Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity) 中 _创建一致的视觉标识您的应用程序_一节。

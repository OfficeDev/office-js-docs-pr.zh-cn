# <a name="highresolutioniconurl-element"></a>HighResolutionIconUrl 元素

指定用于表示插入 UX 中的 Office 加载项和高 DPI 屏幕上的 Office 应用商店的图像的 URL。

**加载项类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>可以包含

[替代](override.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|字符串 (URL)|必需|指定此设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|

## <a name="remarks"></a>备注

对于邮件加载项，图标显示在**文件** > **管理加载项** UI 中。对于内容或任务窗格加载项，图标显示在**插入** > **加载项** UI 中。

图像必须是建议的 64 x 64 像素的分辨率，采用下列一种文件格式：GIF、JPG、PNG、EXIF、BMP 或 TIFF。 有关更多信息，请参阅[在 AppSource 中和 Office 内创建有效列表](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings) 中的 _为应用创建一致的视觉标识_一节。

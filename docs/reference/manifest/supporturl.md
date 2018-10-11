# <a name="supporturl-element"></a>SupportUrl 元素

指定提供外接程序支持信息的页面的 URL。

## <a name="syntax"></a>语法

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a>包含在

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

|  元素 | 必需 | 描述  |
|:-----|:-----|:-----|
|  [替代](override.md)   | No | 指定其他区域设置 URL 的设置 |

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**描述**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必需|指定此设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|

# <a name="appdomains-element"></a>AppDomains 元素

列出了除 Office 加载项用于加载页面的 SourceLocation 元素中指定的域以外的所有域。对于每个其他域，指定 AppDomain 元素。

 **加载项类型：** Content、Task pane、Mail

## <a name="syntax"></a>语法

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a>包含在

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

[AppDomain](appdomain.md)

## <a name="remarks"></a>备注

默认情况下，加载项可以加载与 **SourceLocation** 元素中指定的位置位于同一个域中的任何页面。要加载不与加载项位于同一个域中的页面，请使用 **AppDomains** 和 **AppDomain** 元素来指定域。此元素不能为空。 

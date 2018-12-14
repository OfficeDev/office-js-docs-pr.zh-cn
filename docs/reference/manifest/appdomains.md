# <a name="appdomains-element"></a>AppDomains 元素

列出了除 Office 外接程序用于加载页面的 SourceLocation 元素中指定的域之外的所有域。对于每个其他域，指定 AppDomain 元素。

 **外接程序类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> 每个 **AppDomain** 元素的值都必须包括协议（如 `<AppDomain>https://myappdomain<AppDomain>`）。

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

[AppDomain](appdomain.md)

## <a name="remarks"></a>注释

默认情况下，外接程序可以加载与 [SourceLocation](sourcelocation.md) 元素中指定的位置位于同一个域中的任何页面。 要加载与外接程序位于不同域中的页面，可以使用 **AppDomains** 和 **AppDomain** 元素来指定域。 此元素不能为空。

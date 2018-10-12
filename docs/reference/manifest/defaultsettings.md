# <a name="defaultsettings-element"></a>DefaultSettings 元素

指定内容或任务窗格外接程序的默认源位置和其他默认设置。

**外接程序类型：** 内容、任务窗格

## <a name="syntax"></a>语法

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>包含在

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

|**元素**|**内容**|**邮件**|**任务窗格**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>备注

**DefaultSettings** 元素中的源位置和其他设置仅应用于内容和任务窗格外接程序。对于邮件外接程序，在 [FormSettings](formsettings.md) 元素中指定源文件的默认位置和其他默认设置。


# <a name="sourcelocation-element"></a>SourceLocation 元素

指定 Office 加载项的源文件位置为长介于 1 和 2018 个字符之间的 URL。源位置必须是 HTTPS 地址，而非文件路径。

**加载项类型：** Content、Task pane、Mail

## <a name="syntax"></a>语法

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>包含在

- [DefaultSettings](defaultsettings.md)（内容和任务窗格加载项）
- [FormSettings](formsettings.md)（邮件加载项）
- [ExtensionPoint](extensionpoint.md)（上下文邮件加载项）

## <a name="can-contain"></a>可以包含

[替代](override.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必需|为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置指定此设置的默认值。|

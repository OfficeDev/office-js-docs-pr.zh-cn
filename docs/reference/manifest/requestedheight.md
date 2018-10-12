# <a name="requestedheight-element"></a>RequestedHeight 元素

指定内容加载项或邮件加载项的初始高度（以像素为单位）。 

**加载项类型：** 内容、邮件

## <a name="syntax"></a>语法

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>包含在

- [DefaultSettings](defaultsettings.md)（内容加载项）可以是介于 32 和 1000 之间的值
- [DesktopSettings](desktopsettings.md) 和 [TabletSettings](tabletsettings.md)（邮件加载项）可以是介于 32 和 450 之间的值
- [ExtensionPoint](extensionpoint.md)（上下文邮件加载项）可以是介于 140 和 450（对于 **DetectedEntity** 扩展点）及介于 32 和 450（**CustomPane** 扩展点）之间的值
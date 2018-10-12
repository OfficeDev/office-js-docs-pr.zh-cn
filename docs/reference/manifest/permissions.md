# <a name="permissions-element"></a>Permissions 元素

指定 Office 加载项的 API 访问级别；应基于最低权限原则请求权限。

**加载项类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

对于内容和任务窗格加载项：

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

对于邮件加载项：

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>包含在

[OfficeApp](officeapp.md)

## <a name="remarks"></a>说明

有关详细信息，请参阅[在内容和任务窗格加载项中请求 API 的使用权限](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)和[了解 Outlook 加载项权限](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)。

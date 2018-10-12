 

# <a name="office"></a>Office

Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱集版本要求](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

### <a name="namespaces"></a>Namespaces

[上下文](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。

### <a name="members"></a>成员

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

指定异步调用的结果。

##### <a name="type"></a>类型：

*   String

##### <a name="properties"></a>属性：

|名称| Type| 说明|
|---|---|---|
|`Succeeded`| 字符串|调用成功。|
|`Failed`| String|调用失败。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱集版本要求](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|
####  <a name="coerciontype-string"></a>CoercionType :字符串

指定如何强制由调用方法返回或设置的数据。

##### <a name="type"></a>类型：

*   String

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Html`| String|要求以 HTML 格式返回数据。|
|`Text`| String|要求以文本格式返回数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱集版本要求](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|
####  <a name="sourceproperty-string"></a>SourceProperty :字符串

指定由调用方法返回的数据源。

##### <a name="type"></a>类型：

*   String

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Body`| String|数据源来自邮件的正文。|
|`Subject`| String|数据源来自邮件的主题。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱集版本要求](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|
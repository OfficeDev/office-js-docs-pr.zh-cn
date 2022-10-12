|**权限级别</br>规范名称**|**XML 清单名称**|**Teams 清单名称**|**摘要说明**|
|:-----|:-----|:-----|:-----|
|**限制**|受限|MailboxItem.Restricted.User|允许使用实体，但不允许使用正则表达式。 |
|**读取项**|ReadItem|MailboxItem.Read.User|除了 **限制中允许** 的内容外，它还允许：<ul><li>正则表达式</li><li>Outlook 外接程序 API 读取访问</li><li>获取项属性和回调令牌</li></ul> |
|**读/写项**|ReadWriteItem|MailboxItem.ReadWrite.User|除了 **读取项** 中允许的内容外，它还允许：<ul><li>Outlook 加载项 API 的完全访问权限，但不包括 `makeEwsRequestAsync`</li><li>设置项属性</li></ul> |
|**读/写邮箱**|ReadWriteMailbox|Mailbox.ReadWrite.User|除了 **读/写项目** 中允许的内容外，它还允许：<ul><li>创建、读取、写入项和文件夹</li><li>发送项目</li><li>调用 [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)</li></ul> |

权限在清单中声明。 标记因清单类型而异。

- **XML 清单**：使用该 **\<Permissions\>** 元素。
- **Teams 清单 (预览)**：在“authorization.permissions.resourceSpecific”数组中使用对象的“name”属性。

> [!NOTE]
>
> - 使用附加发送功能的加载项需要补充权限。 使用 XML 清单，可以在 [ExtendedPermissions](/javascript/api/manifest/extendedpermissions) 元素中指定权限。 有关详细信息，请参阅 [Outlook 外接程序中的“实现追加发送](../outlook/append-on-send.md)”。 使用 Teams 清单 (预览) ，可在“authorization.permissions.resourceSpecific”数组的附加对象中使用名称 **Mailbox.AppendOnSend.User** 指定此权限。
> - 使用共享文件夹的加载项需要补充权限。 使用 XML 清单，可以通过将 [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) 元素设置为 `true`指定权限。 有关详细信息，请参阅 [Outlook 外接程序中的“启用共享文件夹”和“共享邮箱”方案](../outlook/delegate-access.md)。 使用 Teams 清单 (预览) ，可在“authorization.permissions.resourceSpecific”数组的附加对象中使用名称 **Mailbox.SharedFolder** 指定此权限。

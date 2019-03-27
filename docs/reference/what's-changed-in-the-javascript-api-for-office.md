# <a name="whats-changed-in-the-javascript-api-for-office"></a>JavaScript API for Office 中的更改内容

JavaScript API for Office 将定期更新新增和更新的对象、方法、属性、事件和枚举，以扩展 Office 外接程序的功能。使用下面的链接可查看新增和更新的 API 成员。

若要使用新的 API 成员开发外接项目，你需要 [在项目中更新适用于 Office 的 JavaScript API 文件](/office/dev/add-ins/develop/update-your-javascript-api-for-office-and-manifest-schema-version)。

若要查看所有 API 成员（包括与之前更新相比未变化的成员），请参阅 [适用于 Office 的 JavaScript API](javascript-api-for-office.md)。

## <a name="new-and-updated-apis"></a>新 API 和更新的 API

### <a name="new-and-updated-objects"></a>新增对象和更新的对象

|**对象**|**说明**|**添加或更新了功能的版本**|
|:-----|:-----|:-----|
|`Item`|更新和新增功能：<br><ul><li><p>`getSelectedDataAsync` 和 `setSelectedDataAsync` 方法支持获取用户所选的内容并将其覆盖到邮件或约会的主题和正文中。</p></li><li><p>`displayReplyAllForm` 和 `displayReplyForm` 方法支持向约会的答复窗体添加附件。</p></li></ul>|Mailbox 1.2|
|`Item`|进行了更新以包括用于创建撰写模式 Outlook 加载项的方法和字段。 |1.1|
|`Binding`|进行了更新以支持 Access 内容加载项中的表绑定。|1.1|
|`Bindings`|进行了更新以支持 Access 内容加载项中的表绑定。|1.1|
|`Body`|进行了添加以便能够在撰写模式 Outlook 加载项中创建和编辑邮件或约会的正文。|1.1|
|`Document`|更新和新增功能： <ul><li><p>支持 Access 内容外接程序中的 <a href="/javascript/api/office/office.document" target="_blank">mode</a>、<a href="/javascript/api/office/office.document#settings" target="_blank">settings</a> 和 <a href="/javascript/api/office/office.document" target="_blank">url</a> 属性。</p></li><li><p>在 PowerPoint 和 Word 外接程序中通过 <a href="/javascript/api/office/office.document#getfileasync-filetype--options--callback-" target="_blank">getFileAsync</a> 方法获取 PDF 格式的文档。</p></li><li><p>在 Excel、PowerPoint 和 Word 外接程序中通过 <a href="/javascript/api/office/office.document#getfilepropertiesasync-options--callback-" target="_blank">getFileProperties</a> 方法获取文件属性。</p></li><li><p>在 Excel 和 Powerpoint 外接程序中通过 <a href="/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-" target="_blank">goToByIdAsync</a> 方法导航到文档中的位置和对象。</p></li><li><p>在 PowerPoint 外接程序中通过 <a href="/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-" target="_blank">getSelectedDataAsync</a> 方法（当指定新的 <span class="keyword">Office.CoercionType.SlideRange</span><a href="/javascript/api/office/office.coerciontype" target="_blank">coercionType</a> 枚举时）获取选定幻灯片的 ID、标题和索引。</p></li></ul>|1.1|
|`Location`|添加了功能以便能够在撰写模式 Outlook 加载项中设置约会的地点。|1.1|
|`Office`|更新了选择方法以支持获取 Access 内容加载项中的绑定。|1.1|
|`Recipients`|添加了功能以便能够在撰写模式下获取和设置邮件或约会的收件人。|1.1|
|`Settings`|进行了更新以支持在 Access 内容加载项中创建自定义设置。|1.1|
|`Subject`|添加了功能以便能够在撰写模式 Outlook 加载项中获取和设置邮件或约会的主题。|1.1|
|`Time`|添加了功能以便能够在撰写模式 Outlook 加载项中获取和设置约会的开始和结束时间。|1.1|

### <a name="new-and-updated-enumerations"></a>新增和更新的枚举

|**对象**|**说明**|**版本**|
|:-----|:-----|:-----|
|`ActiveView`|指定文档活动视图的状态，例如，用户是否可以编辑 document.Added，以便 PowerPoint 的外接程序可以确定用户是否正在查看演示文稿（**幻灯片放映**）或编辑幻灯片。 |1.1|
|`CoercionType`|使用 **Office.CoercionType.SlideRange** 进行了更新，以支持在 PowerPoint 加载项中通过 **getSelectedDataAsync** 方法获取选定幻灯片范围。|1.1|
|`EventType`|进行了更新以包含新的 ActiveViewChanged 事件。|1.1|
|`FileType`|进行了更新以指定 PDF 格式的输出。|1.1|
|`GoToType`|添加了功能以指定要转到的文档位置或对象。|1.1|


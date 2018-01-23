
# <a name="best-practices-for-developing-office-add-ins"></a>开发 Office 外接程序的最佳做法


有效的外接程序提供独特且极具吸引力的功能，采用具有视觉吸引力的方式扩展 Office 应用程序。若要创建出色的外接程序，需为用户提供极具吸引力的首次使用体验、设计一流的 UI 体验和优化外接程序的性能。将本文中描述的最佳实践应用于创建有助于您的用户快速有效地完成其任务的外接程序。

>
  **注意：**生成外接程序时，如果计划将外接程序[发布](../publish/publish.md)到 Office 应用商店，请务必遵循 [Office 应用商店验证策略](https://msdn.microsoft.com/en-us/library/jj220035.aspx)。例如，外接程序必须适用于支持你定义的方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3)以及 [Office 外接程序主机和可用性](https://dev.office.com/add-in-availability)页）。

## <a name="provide-clear-value"></a>提供明确值

- 创建可帮助用户快速、高效地完成任务的外接程序。专注于对 Office 应用程序有用的方案。例如：
 - 使核心创作任务更快、更简单，且中断更少。
 - 在 Office 内启用新方案。
 - 在 Office 主机内嵌入补充服务。
 - 改善 Office 体验来提高工作效率。
- 通过[创建引人入胜的第一次运行体验](#create-an-engaging-first-run-experience)，确保用户能够快速了解你的外接程序的价值。
- 创建 [有效的 Office 商店列表](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)。在标题和说明中明确外接程序的优点。不要依赖于您的品牌来传达您的外接程序的功能。


## <a name="create-an-engaging-first-run-experience"></a>创建极具吸引力的首次运行体验



- 要用具有高可用性和直观性的首次体验吸引新用户。请注意，用户从商店下载外接程序之后，仍可决定是使用还是放弃该外接程序。

 - 明确用户与您的外接程序交互所需执行的步骤。使用视频、泡沫垫、分页面板或其他资源来吸引用户。

 - 在启动时强调您的外接程序的价值主张，而不只是让用户登录。

 - 提供用以指导用户的教学 UI，并使您的 UI 富有个性化。

    ![显示没有入门步骤的外接程序旁边具有入门步骤的外接程序任务窗格的屏幕截图](../images/586202ad-333b-417c-ad31-cc6eb952b239.png)

  - 如果内容外接程序绑定到用户文档中的数据，请将那些用于向用户显示要使用的数据格式的示例数据或模板包含在内。

    ![显示没有数据的内容外接程序旁边具有数据的内容外接程序的屏幕截图](../images/7de2215f-ccef-4f82-aa9d-babcbddae0c6.png)

- 提供 [免费试用版](https://msdn.microsoft.com/en-us/library/dn456317.aspx#Anchor_1)。如果外接程序需要订阅，那么让某些功能在不订阅的情况下也可使用。

- 让注册非常简单。预先填充某些信息（如电子邮件、显示名称），并跳过电子邮件验证。

- 避免弹出窗口。如果您必须使用它们，请引导用户启用您的弹出窗口。

- 使用[单一登录 (SSO) 身份验证](../outlook/authenticate-a-user-with-an-identity-token.md)。

对于说明你在开发首次运行体验时可以应用的模式的模板，请参阅[适用于 Office 外接程序的 UX 设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)。

## <a name="use-add-in-commands"></a>使用外接程序命令

- 使用外接程序命令为你的外接程序提供相关 UI 入口点。有关详细信息（包括设计最佳做法），请参阅[外接程序命令](../design/add-in-commands.md)。

## <a name="apply-ux-design-principles"></a>应用 UX 设计原则

- 确保你的外接程序的外观和功能很好地补充了 Office 体验。使用 [Office UI Fabric](https://dev.office.com/fabric)。

- 支持内容胜过支持部件版式。避免使用对用户体验毫无价值的不必要的 UI 元素。

- 保持用户处于可控状态。确保用户了解重要的决定，并且可以轻松地倒退外接程序执行的操作。

- 使用品牌唤起用户的信任感和亲切感。但不要过度使用品牌或向用户做广告推销。

- 避免滚动。优化为 1366 x 768 分辨率。

- 不包含未授权的图像。

- 在外接程序中使用 [清楚而简单的语言](../design/voice-guidelines.md)。

- 考虑 [可访问性](../design/accessibility-guidelines.md) - 方便所有用户与其进行交互，并适应屏幕读取器等辅助技术。

- 针对所有平台和输入方法（包括鼠标/键盘和 [触摸](#optimize-for-touch)）的设计。确保您的 UI 可响应不同的外观设置。

对于应用你在开发外接程序时可以使用和自定义的设计原则的模板，请参阅[适用于 Office 外接程序的 UX 设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)。

### <a name="optimize-for-touch"></a>触摸优化



- 使用 [Context.touchEnabled](http://dev.office.com/reference/add-ins/shared/office.context.touchenabled) 属性检测运行你的外接程序的主机应用程序是否已启用触控。

     >**注意**  该属性在 Outlook 中不受支持。
- 确保所有控件都相应符合触控交互的尺寸大小。例如，按钮有足够大的触摸目标，且输入框要足够大，方便用户输入。

- 不依赖于诸如悬停或用鼠标右键单击等非触摸式输入方法。

- 确保外接程序可以在纵向和横向模式中正常工作。请注意在触控设备上，外接程序的一部分可能通过软键盘隐藏。

- 通过使用 [旁加载](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)在实际的设备上测试外接程序。


 >**注释**  如果您将 [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) 用于您的设计元素，则许多这些元素都能得到满足。


## <a name="optimize-and-monitor-add-in-performance"></a>优化和监视外接程序性能



- 创建快速 UI 响应的感觉。外接程序的加载时间应在 500 毫秒以内。

- 确保所有用户交互响应时长都在一秒内。

-  为长时间运行的操作提供加载指示器。

- 将 CDN 用于主机图像、资源和公用库。尽可能地从一个位置进行加载。

- 请按照标准 Web 实践来优化您的网页。在生产中，仅使用库的缩小版本。仅加载所需的资源，并优化加载资源的方式。

- 如果操作需要执行时间，请为用户提供反馈。请注意下表中列出的阈值。另请参阅 [Office 外接程序的资源限制和性能优化](../develop/resource-limits-and-performance-optimization.md)


|**交互类**|**目标**|**上限**|**人类感知**|
|:-----|:-----|:-----|:-----|
|即时|<=50 毫秒|100 毫秒|没有明显的延迟。|
|快速|50-100 毫秒|200 毫秒|最小限度的明显延迟。不需要反馈。|
|典型|100-300 毫秒|500 毫秒|较快，但不够快，不能称之为快速。不需要反馈。|
|快速响应|300-500 毫秒|1 秒|不快，但仍然感觉反应灵敏。不需要反馈。|
|连续|> 500 毫秒|5 秒|中等等待时间，不再感觉反应灵敏。可能需要反馈。|
|受限|> 500 毫秒|10 秒|较长，但不足以执行其他操作。可能需要反馈。|
|扩展|> 500 毫秒|> 10 秒|长到足以在等待时执行其他操作。可能需要反馈。|
|长时间运行|> 5 毫秒|> 1 分钟|用户当然可以执行其他操作。|
- 监视您的服务运行状况，并使用遥测监视用户的成功。


## <a name="market-your-add-in"></a>推销您的外接程序



- 将外接程序发布到 [Office 应用商店](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)并且从您的网站 [对其进行推广](http://msdn.microsoft.com/library/b19e21f8-76f5-44e1-9971-bef79cad4c71%28Office.15%29.aspx)。创建 [高效 Office 应用商店列表](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)。

- 使用简洁且富有描述性的外接程序标题。包括但不能超过 128 个字符。

- 为您的外接程序撰写简短且富有吸引力的描述。回答"此外接程序解决哪些问题？"这一问题。

- 在您的标题和说明中传达外接程序的价值主张。不要依赖于您的品牌。

- 创建一个网站用以帮助用户查找和使用外接程序。


# <a name="accessibility-guidelines-for-office-add-ins"></a>Office 外接程序的辅助功能准则

在设计和开发 Office 外接程序时，你将需要确保所有潜在用户和客户都能够成功使用你的外接程序。请应用以下准则，确保所有目标用户都能访问你的解决方案。

##<a name="design-for-multiple-input-methods"></a>针对多种输入方式的设计

- 确保用户可以仅通过键盘执行操作。用户应该能够使用 Tab 键和箭头键组合移动到页面上的所有可操作元素。
- 在移动设备上，当用户通过触摸操作某个控件时，设备应该提供有用的音频反馈。
- 为所有交互式控件提供有用的标签。 

##<a name="make-your-add-in-easy-to-use"></a>使你的外接程序易于使用

- 不依赖于单个属性（例如颜色、大小、形状、位置、方向或声音）在 UI 中传达含义。
- 避免对上下文的意外更改，例如在用户未操作的情况下将焦点移到其他 UI 元素。
- 提供验证、确认或撤消所有绑定操作的方法。
- 提供暂停或停止媒体（例如音频和视频）的方法。
- 不对用户操作施加时间限制。

##<a name="make-your-add-in-easy-to-see"></a>使你的外接程序易于查看

- 避免对颜色的意外更改。
- 提供有意义且及时的信息，以描述 UI 元素、标题和标头、输入和错误。确保控件名称恰当地描述了控件的意图。
- 遵循针对颜色对比度的[标准准则](http://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html)。

##<a name="account-for-assistive-technologies"></a>对辅助技术负责

- 避免使用会干扰辅助技术的功能，例如视觉、音频或其他交互。
- 请勿以图像格式提供文本。屏幕阅读器无法读取图像中的文本。
- 为用户提供调整或静音所有音频源的方法。
- 为用户提供打开字幕或音频说明与音频源的方法。
- 提供用于警示用户的声音替代方法，例如视觉提示或振动。

##<a name="accessibility-resources"></a>辅助功能资源

- [Web 内容辅助功能准则 (WCAG) 2.0](http://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [向非 Web 信息和通信技术 (WCAG2ICT) 应用 WCAG 2.0 的指南](http://www.w3.org/TR/wcag2ict/)
- [关于信息和通信技术 (ICT) 的辅助功能要求的欧洲标准](http://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf)



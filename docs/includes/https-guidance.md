强烈建议对加载项使用 HTTPS 终结点（尽管无需在所有加载项方案中都严格遵循此要求）。 不受 SSL (HTTPS) 保护的加载项会在使用期间生成不安全的内容警告和错误。如果计划在 Office 网页版中运行加载项或将加载项发布到 AppSource，加载项必须受 SSL 保护。如果加载项访问外部数据和服务，它应受 SSL 保护，以保护传输中的数据。自签名证书可用于开发和测试，但前提是证书在本地计算机上受信任。


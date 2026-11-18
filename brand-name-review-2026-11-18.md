# 商家名称审核修复（2026-11-18）

Google Ads 拒登 `DeFlame Tech Co. Limited` 的原因是“商家信息 - 名称的显眼程度”。原落地页只在页脚出现公司名称，首屏没有明确的公司与产品关系。

本次静态稿在首屏品牌区域增加了可见的归属说明：

`DeFlame Tech Co. Limited · PenSignal`

修改文件：

- `index.html`
- `preview.html`
- `styles.css`

需要先把 `index.html` 的对应改动接入并部署到正式路由 `https://pensignal.deflametech.com/en/bulk-seo-articles`，确认首屏能看到完整公司名称，再回到 Google Ads 对现有商家名称资产发起重新审核。只在本地稿中修改，或在页面仍只显示页脚名称时重新提交，仍可能被拒。

建议复核：

1. 无 Cookie 同意也能看到首屏公司名称。
2. 桌面和手机宽度下公司名称完整显示，不被截断。
3. 页面品牌、已认证广告主名称和商家名称资产保持一致。

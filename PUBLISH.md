# 发布流程（npm）

> 包名：`@excel-preview/core`（工作区：`packages/core`）
> 发布统一走 GitHub Actions CI（支持 npm provenance 可信构建），**不要**在本地直接 `npm publish`。

---

## 一、为什么走 CI 而非本地发布

- 本地 `npm publish` 会因账号开启 2FA 触发 `EOTP`（要求一次性验证码），且登录会话/验证码获取繁琐。
- CI 通过 GitHub Actions 的 OIDC 令牌发布，自动带上 `provenance`（来源证明），更安全且无需人工交互。
- 失败可回看 CI 日志，便于排查。

## 二、发布前检查

```bash
# 1. 确认版本号已递增（packages/core/package.json 的 version）
# 2. 本地构建 + 测试 + 打包检查
pnpm install
pnpm --filter @excel-preview/core build
pnpm --filter @excel-preview/core test
pnpm --filter @excel-preview/core pack --dry-run   # 确认产物包含 dist/
```

## 三、提交与打标签

```bash
git add -A
git commit -m "chore: bump version to 0.1.x"        # 版本号与 package.json 一致
git push origin main

git tag v0.1.x                                       # 与版本号一致
git push origin v0.1.x
```

> 打标签非必需（`workflow_dispatch` 直接跑 main 即可），但建议保持 tag 与发布版本一一对应。

## 四、触发发布（二选一）

### 方式 A：GitHub 网页手动触发

1. 打开仓库 Actions 页面：https://github.com/zhanghongyi1/excel-viewer/actions
2. 左侧选中 **Publish Package to npm** 工作流
3. 点 **Run workflow** → 分支选 `main` → 点绿色按钮

### 方式 B：gh CLI 触发（自动化）

```bash
# 首次需要登录
gh auth login        # 选 GitHub.com → HTTPS → 浏览器授权（输入终端显示的设备码）

# 触发并等待结果
gh workflow run publish.yml --ref main
gh run watch --exit-status
```

## 五、验证发布结果

```bash
npm view @excel-preview/core version   # 输出应为刚发布的版本号
npm view @excel-preview/core versions  # 查看全部已发布版本
```

成功后会收到 npm 官方邮件通知（包含 CI 运行链接与 shasum）。

## 六、CI 工作流说明

文件：`.github/workflows/publish.yml`

- 触发条件：`release: published` 或 `workflow_dispatch`（手动）
- 权限：`contents: read` + `id-token: write`（供 provenance 使用）
- 步骤：checkout → pnpm 8.15 → Node 22 → 安装依赖 → 构建 → 测试 → `npm publish --provenance --access public`
- 关键：发布前先 `npm install -g npm@latest`，确保 npm 版本支持 provenance。

## 七、常见问题

| 现象 | 原因 | 处理 |
|------|------|------|
| 本地发布报 `EOTP` | 账号开启 2FA | 改用 CI 发布，见「四」 |
| `gh` 未登录 | 未执行 `gh auth login` | 先登录再触发 |
| CI 发布失败：版本已存在 | package.json 版本未递增 | 递增版本号并重新提交 |
| 收不到 npm 通知邮件 | 触发后邮箱延迟/进垃圾箱 | 以 `npm view` 结果为准 |

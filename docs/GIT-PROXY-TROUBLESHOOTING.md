# Git 代理连接失败的诊断与处理

验证日期：2026-10-29（Asia/Shanghai）。
适用环境：当前 Codex 云端环境，仓库位于 `/workspace/DocumentModify`。

## 问题与原因

执行 Git 远端读取时出现：

```text
fatal: unable to access 'https://github.com/HolyHighY/DocumentModify.git/':
Failed to connect to proxy port 8080 after 0 ms: Could not connect to server
```

本次实测确认：默认命令沙箱阻止了连接代理的网络请求。底层套接字检查返回 `PermissionError: [Errno 1] Operation not permitted`。平台报告网络策略已生效，且已允许 `github.com`；同一 Git 命令在获准的网络执行权限下成功。

因此，本次错误无需通过更换代理、添加 GitHub Token 或扩大域名白名单解决。仅凭上述 Git 报错，不能将其他环境中的失败也判定为同一原因。

## 已验证的处理方法

1. 保留平台注入的代理、认证和 TLS 校验设置。
2. 在现有仓库执行只读检查：

   ```bash
   cd /workspace/DocumentModify
   git ls-remote origin HEAD
   ```

3. 如果默认沙箱阻止网络连接，由 Codex 使用工具支持的 `require_escalated` 执行权限重新运行该命令，并遵守审批结果。它是工具参数，不是终端命令、Git 选项或环境变量。

   本次成功调用的参数如下：

   ```json
   {
     "cmd": "git ls-remote origin HEAD",
     "workdir": "/workspace/DocumentModify",
     "sandbox_permissions": "require_escalated",
     "justification": "需要在允许网络的执行上下文中进行只读 Git 检查，以确认是否是命令沙箱阻止了连接现有云端代理。",
     "prefix_rule": ["git", "ls-remote"]
   }
   ```

4. 确认命令退出码为 `0`，并返回提交 SHA 与 `HEAD`。本次结果为：

   ```text
   dd9aadb878e411260cabb7f6ee337183cba7a113    HEAD
   ```

   该 SHA 仅是验证时的结果；远端更新后可以不同。

## 若仍然失败

- 若工具审批拒绝执行，记录拒绝原因，通过支持的审批流程处理，不绕过限制。
- 若返回域名策略拒绝，检查环境网络设置，按实际需要添加缺失域名。
- 若返回认证或授权错误，先检查已有平台认证与仓库访问权限，再确定是否需要补充配置。
- 若在获准的执行上下文中仍无法连接代理，继续排查云端代理服务与平台网络状态；不能沿用本次“仅沙箱限制”的结论。

不应取消代理绕过平台网络路径，也不应关闭 TLS 校验或在日志、文档中记录凭据。

## 验证范围与保存状态

已验证当前仓库的远端只读访问；未验证推送权限或 GitHub API 访问。

本机操作记录保存在 `.local/environments/cloud-onboarding/LOCAL-WORKFLOW.md`，该文件被 Git 忽略，不随本仓库分发。修正后的处理方法已保存到环境配置的 `start_skill` 草稿；保存草稿不等于发布，需在环境设置中检阅、保存并发布后供后续环境使用。本文是排查方法的仓库副本；本机核验记录仍保留在 `.local/` 中。

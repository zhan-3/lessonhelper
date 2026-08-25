# 教务系统接口探索器

第一阶段只负责发现与毕业条件有关的数据接口，不计算成绩，也不会执行选课、退课、保存或提交操作。

## 安装浏览器

默认使用 Playwright 管理的 Chromium：

```powershell
uv run playwright install chromium
```

## 开始探索

```powershell
uv run python -m course_progress explore
```

浏览器打开后：

1. 手动完成 CAS / WebVPN 登录。
2. 手动访问培养方案、已修课程、学业完成情况等只读页面。
3. 登录成功后脚本会自动寻找并访问只读入口。
4. 自动探索结束后查看终端输出的 `candidates.json` 路径。

自动导航只点击或访问包含培养方案、学业完成、毕业要求、课程、学分等关键词的同源入口，并跳过选课、退课、提交、保存、删除等操作。默认最多访问 12 个页面，可用 `--max-pages` 调整。

如果 Chromium 无法正常完成 CAS 登录，可改用本机 Chrome：

```powershell
uv run python -m course_progress explore --browser chrome
```

## 数据目录

登录状态和捕获数据位于 `.private/course-progress/`：

```text
.private/course-progress/
├── browser-profile/
└── captures/
    └── <时间戳>/
        ├── index.jsonl
        ├── candidates.json
        ├── raw/
        └── sanitized/
```

- `browser-profile/` 保存探索器和采集器共用的登录状态。
- 如果目录中仍有旧的 `collector-profile/` 或 `explorer-profile/`，程序会优先复用它们，不会合并两个浏览器目录。
- `raw/` 可能包含个人课程数据，只能留在本机。
- `sanitized/` 删除常见 Token、Ticket、学号、姓名和联系方式。
- `candidates.json` 按培养方案、毕业、课程和学分关键词评分排序。

整个 `.private/` 目录已被 Git 忽略，不能打包或分享给其他用户。

## 采集课程并计算进度

接口探索完成后，使用独立的只读采集命令：

```powershell
uv run python -m course_progress collect
```

命令会打开可视浏览器。完成教务系统登录并进入系统后，它会自动：

1. 打开期末成绩查询；
2. 动态读取所有学期；
3. 逐学期查询全部成绩记录，并在本地根据最终成绩判断是否通过；
4. 遍历每个学期的全部分页；
5. 排除必修、去重并按指南基线计算进度；
6. 将结果保存到 `.private/course-progress/progress-report.json`。

报告不保存具体成绩、姓名、学号或登录 Cookie。采集过程中每完成一个分页，会把不含具体成绩的课程完成事实保存到 `collection-checkpoint.json`；登录失效时会等待人工重新认证，并从当前分页继续。某个学期、分页或登录状态最终失效时，报告会标记数据不完整，不会把失败误算成零学分。

可选参数：

```powershell
uv run python -m course_progress collect --help
uv run python -m course_progress collect --browser chrome
uv run python -m course_progress collect --login-timeout-seconds 600
uv run python -m course_progress collect --page-size 20
```

## 可选参数

```powershell
uv run python -m course_progress explore --help
uv run python -m course_progress explore --max-response-mb 10
uv run python -m course_progress explore --url "https://example.edu/portal"
uv run python -m course_progress explore --login-timeout-seconds 600
```

## 隐私说明

探索器不会要求或保存明文账号密码。首次登录始终由用户在可视浏览器中完成。不要分享 `browser-profile/`、旧的 profile 目录、`collection-checkpoint.json`、`raw/` 或任何未检查的网络捕获文件。程序不会主动发送保活请求，也不会伪造或延长学校认证令牌的有效期。

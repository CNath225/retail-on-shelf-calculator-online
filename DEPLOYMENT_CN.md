# Calculator Online 部署步骤

这个项目用于部署：

Retail On-shelf Rate Calculator Online by CodeNATHAN

## 你需要准备

- 一个 GitHub 账号
- 一个 Streamlit Community Cloud 账号
- 这个项目文件夹里的代码

## GitHub 上传

推荐 repo 名：

```text
retail-on-shelf-calculator-online
```

上传到 GitHub 时，不要上传这些内容：

- `.venv/`
- `.runtime/`
- 任何真实 raw data Excel
- 任何客户/门店敏感数据

这些已经写在 `.gitignore` 里。

## Streamlit Community Cloud 设置

创建新 app 时：

```text
Repository: retail-on-shelf-calculator-online
Branch: main
Main file path: app.py
```

部署后，Streamlit 会读取：

```text
requirements.txt
runtime.txt
.streamlit/config.toml
```

## 线上使用方法

1. 上传 Repsly raw export。
2. 上传 Range / Template workbook。
3. 如果 range workbook 和 report template 是同一个文件，保持 `Use range file as template` 勾选。
4. 检查 Month、Report Column、Compare To。
5. 点击 Generate Report。
6. 下载生成的 Excel。

## 更新网站

以后改代码后：

```bash
git add .
git commit -m "Update calculator online"
git push
```

Streamlit Cloud 会从 GitHub 重新部署。

## 注意

这是无数据库的在线计算器。上传文件只用于当次计算，不做长期业务数据管理。需要长期保存 master data、range、SKU、account 的功能，应该放在 Automation 项目里。

# 用 Oxi 为 Cloudreve 生成缩略图

[English](README.md)

[Cloudreve](https://github.com/cloudreve/cloudreve) 生成文档缩略图时，会运行
LibreOffice：`soffice --headless --convert-to png --outdir DIR FILE`。
Oxi 的 `oxidocs` 能接受同样的命令行，按 Word、PowerPoint、Excel 的排版方式绘制
**.docx、.pptx、.xlsx** 的第一页；其他文件仍交给 LibreOffice。
Cloudreve 本身无需任何改动：不需要插件，也不需要修改设置。

在容器中处理一个 Word 文档，oxidocs 用时 0.13 秒、内存 100 MB；
LibreOffice 用时 2.0 秒、内存 228 MB。同时生成 15 个缩略图也全部成功。
LibreOffice 因共用配置文件，Cloudreve 文档建议只开 1～2 个任务并发，
oxidocs 没有这个限制。

## 使用 Docker

```bash
docker compose up -d --build      # 在本目录执行
```

[Dockerfile](Dockerfile) 以官方 `cloudreve/cloudreve` 镜像为基础，增加：

- oxidocs：固定版本和 SHA-256 的静态构建，放在 `/opt/oxidocs`
- `/usr/local/bin/soffice` → oxidocs。它在 PATH 中排在 `/usr/bin` 之前，
  因此 Cloudreve 默认设置所调用的 `soffice` 就是它。
- `OXIDOCS_SOFFICE_FALLBACK=/usr/bin/soffice`：.doc、.odt、.pdf、.txt 等文件
  照旧交给 LibreOffice。

已有的安装：构建此镜像后，挂载原来的数据卷运行即可。如果曾在
*设置 → 媒体处理* 中把 LibreOffice 路径从 `soffice` 改成了完整路径，
请改为 `/usr/local/bin/soffice`。

## 不使用 Docker

从 [Releases](https://github.com/Ryujiyasu/oxi/releases)（`oxidocs-v*`）下载
`oxidocs-<arch>-unknown-linux-musl.tar.gz`，解压到单独的目录，在该目录运行一次
`oxidocs fonts-index`（它会把自带的字体写到程序旁边），然后链接为 `soffice`：

```bash
ln -s /opt/oxidocs/oxidocs /usr/local/bin/soffice
export OXIDOCS_SOFFICE_FALLBACK=/usr/bin/soffice   # 写入 Cloudreve 的运行环境
```

也可以把 Cloudreve 的 LibreOffice 路径直接指向这个链接。若未安装 LibreOffice，
请不要设置 fallback，并把 LibreOffice 的扩展名列表改为
`docx,docm,dotx,pptx,ppsx,potx,xlsx,xlsm,xltx`。否则其他扩展名会失败，
而 Cloudreve 对缩略图失败过一次的文件不会再尝试。

## 验证

`e2e/run.sh` 会构建镜像、启动 Cloudreve，上传 .docx、.pptx、.xlsx 和 .txt，
并获取它们的缩略图。LibreOffice 经由一个记录调用的包装脚本运行。
三个 Office 文件中任何一个交给了 LibreOffice，或者文本文件没有交给它，
检查都会失败。

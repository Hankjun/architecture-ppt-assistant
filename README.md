# 建筑汇报 PPT 自动助手

围绕建筑汇报的资料整理、结构化输入和人工确认，探索从需求到预览的可控交付流程。

**[项目介绍](https://hankjun.github.io/architecture-ppt-assistant/)** · [汇报预览](https://hankjun.github.io/architecture-ppt-assistant/preview.html) · [客户咨询原型](https://hankjun.github.io/architecture-ppt-assistant/consult.html)

Architecture presentation workflow prototype with a Python HTML preview generator.

## 为什么做这个项目

建筑汇报往往需要处理零散资料、专业结构和反复修改。本项目把工作拆成：客户咨询 → 资料盘点 → 风格确认 → 框架生成 → HTML 预览 → 按页反馈 → 最终交付。

公开仓库提供展示页面、咨询页原型和基于结构化数据的预览生成器，便于讨论与验证交付流程。

## 已实现的公开能力

- 项目说明页与客户咨询交互原型。
- Python 生成 HTML 多页预览，支持深浅主题。
- 封面、目录、内容、双栏、对比、图库占位、引用和结束页共 8 种模板。
- 浏览器打印入口，可通过打印功能保存 PDF。
- 实验性 PPTX 导出：`python-pptx` 路径目前提取标题生成基础幻灯片，不保留完整内容、图片和 HTML 布局。

公开代码尚未包含完整的模型调用、建筑资料检索系统或客户反馈后台。页面提及的素材库与业务数据不随仓库提供。

## 本地预览

需要 Git 和 Python 3；HTML 生成功能仅使用 Python 标准库。

```bash
git clone https://github.com/Hankjun/architecture-ppt-assistant.git
cd architecture-ppt-assistant
python3 -m http.server 8877 --bind 127.0.0.1
```

打开 [http://localhost:8877](http://localhost:8877)。

## 生成自己的预览

在仓库目录运行：

```bash
python3 generate_preview.py --out preview-local.html
python3 generate_preview.py --theme light --out preview-local-light.html
python3 generate_preview.py --list-types
```

生成的文件可直接打开，也可通过本地 HTTP 服务查看。

自定义数据示例，保存为 `project.json`：

```json
{
  "title": "建筑汇报示例",
  "pages": [
    {
      "type": "cover",
      "title": "建筑汇报示例",
      "subtitle": "概念方案",
      "meta": "演示数据"
    },
    {
      "type": "end",
      "title": "感谢聆听",
      "subtitle": "敬请指正"
    }
  ]
}
```

```bash
python3 generate_preview.py --json project.json --out preview-local.html
```

其他模板的字段可参考 `generate_preview.py` 中的 `DEFAULT_SECTIONS`。图库模板当前输出色块占位。生成器直接将输入插入 HTML，请使用可信的本地数据。

## 项目结构

| 文件 | 用途 |
| --- | --- |
| `index.html` | 项目背景与工作流介绍 |
| `consult.html` | 客户咨询页原型 |
| `generate_preview.py` | HTML 预览生成与实验性 PPTX 导出 |
| `preview.html` | 深色预览示例 |
| `preview-light.html` | 浅色预览示例 |

## 下一步

优先完善输入校验与图片支持，再推进按页反馈、素材检索和完整导出。以一份示例输入和可核对的输出展示进展。

## 反馈

欢迎通过 [Issues](https://github.com/Hankjun/architecture-ppt-assistant/issues) 提交模板、预览和交付流程建议。请附上使用步骤与预期效果。

维护者：[Hank / 汉克](https://github.com/Hankjun)。

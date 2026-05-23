---
name: ppt-skill
description: >
  小红书清新风格PPT全流程助手。支持从主题生成PPT、编辑已有.pptx文件、基于模板套用内容。
  触发词包括："做PPT""生成PPT""帮我做一个PPT""做个演示文稿""编辑这个PPT""修改PPT"
  "套模板""基于模板生成PPT"等。也适用于用户说"做个关于XX的PPT""把这个内容做成PPT"
  "帮我排版成演示文稿"等场景。默认输出小红书清新风格.pptx文件，如需其他风格可指定。
  当用户需要生成、创建、编辑、修改PowerPoint演示文稿时使用此skill。
---

# PPT Skill — 小红书清新风演示文稿全流程助手

## 概述

本 Skill 覆盖 PPT 全生命周期：从零生成、编辑已有文件、模板套用。
默认风格为小红书清新风（暖调、大面积留白、柔和配色），输出标准 .pptx 文件。

## 工作流选择

收到用户需求后，按以下决策树选择工作流：

```
用户需求
│
├─ "帮我做一个关于XX的PPT" / "生成PPT" 等创建类
│   └─ 工作流 A：从主题生成
│
├─ "修改这个PPT" / "把第X页改成..." / "删掉第X页" 等编辑类
│   └─ 工作流 B：编辑已有文件
│
├─ "基于这个模板做PPT" / "套用模板" 等模板类
│   └─ 工作流 C：模板套用
│
└─ "把这个 Markdown / 文章 / 笔记做成PPT" 等转换类
    └─ 工作流 A（需先提取结构化内容）
```

---

## 工作流 A：从主题生成 PPT

### 第一步：需求澄清

向用户确认以下信息（已有信息则跳过）：

1. **主题**：PPT 要讲什么？
2. **页数**：大概多少页？（默认 8-12 页）
3. **受众**：给谁看？（默认：小红书读者 / 普通大众）
4. **素材**：是否有文字素材、图片、数据？
5. **主题色**：从 `references/themes.md` 选择，未指定默认蜜桃粉
6. **硬约束**：是否有品牌色、logo、必须包含的内容？

### 第二步：生成大纲

根据澄清结果，产出结构化大纲，格式如下：

```json
{
  "title": "PPT标题",
  "subtitle": "副标题（可选）",
  "theme": "peach-pink",
  "slides": [
    {"layout": "cover", ...},
    {"layout": "toc", ...},
    {"layout": "title_content", "title": "...", "content": ["要点1", "要点2"]},
    ...
    {"layout": "ending", ...}
  ]
}
```

展示大纲给用户确认。确认后再执行生成。

### 第三步：执行生成

使用 `python_executor` 或 `shell_executor` 运行：

```bash
python scripts/generate_pptx.py --output <output_path> --content <content_json_path>
```

- `<output_path>`：输出 .pptx 文件路径，默认写到用户偏好目录 `D:\marvis`
- `<content_json_path>`：第二步生成的大纲 JSON 文件路径

如果包含图片，需先将图片路径写入 JSON 的对应字段。

### 第四步：自检与交付

对照 `references/checklist.md` 进行 P0 和 P1 级检查。

用 `yyb-product` 声明产出物。

---

## 工作流 B：编辑已有 PPT

### 操作类型

| 用户表述 | 对应操作 |
|---------|---------|
| "把第X页的XXX改成YYY" | 文本替换 |
| "删掉第X页" | 删除幻灯片 |
| "把第X页移到第Y页前面" | 调整顺序 |
| "复制第X页" | 复制幻灯片 |
| "把标题加大/变色/居中" | 应用样式 |

### 执行方式

1. 确认目标 PPT 文件路径
2. 构造编辑参数 JSON
3. 运行脚本：

```bash
python scripts/edit_pptx.py --file <源文件> --action <操作类型> --params '<JSON参数>' --output <输出文件>
```

操作类型与参数格式详见 `scripts/edit_pptx.py` 的参数说明。

**注意**：编辑操作默认输出到新文件（加 `_edited` 后缀），避免覆盖原文件。

---

## 工作流 C：模板套用

当用户提供模板 .pptx 文件并有内容要填充时使用。

### 步骤

1. 确认模板路径和内容来源
2. 将内容组织为 JSON（格式见 `scripts/apply_template.py` 的说明）
3. 运行脚本：

```bash
python scripts/apply_template.py --template <模板路径> --content <内容JSON> --output <输出路径>
```

模板套用支持三种填充方式：按 placeholder index、按形状名称、按位置顺序。
详见 `scripts/apply_template.py` 的 `fill_slide` 函数。

---

## 与已有 PPT 工具的协作

本 Skill 与系统已有的 `Skywork-ppt` (use_skill) 并行可用：

- **本 Skill (ppt-skill)**：python-pptx 本地操作，适合精细化控制、编辑、模板套用
- **Skywork-ppt**：远程 API 生成，适合快速从主题生成完整 PPT，适合需要 AI 内容创作能力的场景

**调度建议**：
- 需要从零生成内容丰富的 PPT → 优先用 `ppt-skill` 工作流 A（本地生成，可控性强）
- 编辑已有 .pptx 文件 → 必须用 `ppt-skill` 工作流 B
- 需要 AI 自动配图 + 生成 → 可先用 `ppt-skill` 生成内容，再按需调用其他图片生成能力

---

## 主题与布局

- 主题配色：见 `references/themes.md`（5套小红书清新风主题）
- 页面布局：见 `references/layouts.md`（10种布局骨架）
- 质量检查：见 `references/checklist.md`（P0-P3分级清单）

**配色铁律**：只允许从 themes.md 的 5 套预设中选择，不允许自定义 hex 值。

---

## 常见错误与避坑

1. **不要忘记封面和结束页** — 工作流 A 的脚本会自动添加，但手动构造 JSON 时别遗漏
2. **图片路径必须是绝对路径** — 相对路径会导致图片缺失
3. **连续同布局不超过 3 页** — 破坏视觉节奏
4. **文字不溢出** — 单页正文建议不超过 150 字
5. **编辑操作默认生成新文件** — 不要直接覆盖用户原文件
# PPT Skill — 小红书清新风演示文稿全流程助手

> 🎨 一个为 Marvis / Claude Agent 设计的 PPT 生成与编辑 Skill，输出小红书清新风格 .pptx 文件

## 功能特性

- **全流程覆盖**：从零生成、编辑已有、模板套用三条工作流
- **小红书清新风**：5套预设主题色（蜜桃粉/薄荷绿/天空蓝/奶油黄/薰衣草紫）
- **本地优先**：基于 python-pptx 库，无外部 API 依赖
- **质量自检**：P0-P3 四级检查清单，确保产出可用
- **协作友好**：与 Skywork-ppt 等远程 API 能力互补

## 快速开始

### 安装方式

1. **本地安装**（推荐）：
   ```bash
   # 将以下指令发给 Marvis / Claude Agent：
   帮我安装 ppt-skill。请把 https://github.com/<你的用户名>/ppt-skill 克隆到 ~/.claude/skills/ppt-skill，安装完成后检查 SKILL.md、assets/、references/ 是否存在。
   ```

2. **手动安装**：
   ```bash
   git clone https://github.com/<你的用户名>/ppt-skill.git ~/.claude/skills/ppt-skill
   ```

### 使用示例

**生成 PPT**：
```
帮我做一个关于AI效率工具的PPT，大概8页，用蜜桃粉主题
```

**编辑 PPT**：
```
把这个PPT的第3页标题改成"核心优势"，把第5页删掉
```

**模板套用**：
```
基于这个模板文件，把这份JSON内容填充进去生成PPT
```

## 工作流

### 工作流 A：从主题生成
1. 需求澄清（主题/页数/受众/素材/主题色）
2. 生成大纲并确认
3. 执行 python-pptx 生成
4. 质量自检后交付

### 工作流 B：编辑已有文件
- 文本替换
- 增删幻灯片
- 调整顺序
- 应用样式

### 工作流 C：模板套用
- 按 placeholder index 匹配
- 按形状名称匹配
- 按位置顺序填充

## 文件结构

```
ppt-skill/
├── SKILL.md              # Skill 主文件：工作流、决策树、协作指南
├── README.md             # 本文件
├── scripts/
│   ├── generate_pptx.py   # 从主题生成 PPT
│   ├── edit_pptx.py       # 编辑已有 PPT
│   └── apply_template.py  # 模板套用内容
├── references/
│   ├── themes.md          # 5套小红书清新风主题配色
│   ├── layouts.md         # 10种页面布局骨架
│   └── checklist.md       # P0-P3 质量检查清单
└── assets/
    └── (预留模板文件位置)
```

## 主题配色（只允许从预设选择）

| 主题 | 十六进制 | 适合场景 |
|------|---------|----------|
| 蜜桃粉 (peach-pink) | `#FF7099` | 生活方式、美妆、穿搭、女性向 |
| 薄荷绿 (mint-green) | `#64C9B4` | 健康、自然、环保、科普 |
| 天空蓝 (sky-blue) | `#5B9BD5` | 科技、效率工具、职场、学习 |
| 奶油黄 (cream-yellow) | `#F0B27A` | 美食、旅行、家居、手工 |
| 薰衣草紫 (lavender) | `#B39DDB` | 读书、电影、艺术、情感 |

**配色铁律**：只允许从以上 5 套预设中选择，禁止自定义 hex 值，保护美学一致性。

## 布局类型

1. **封面 (Cover)** - 演示文稿第一页
2. **目录 (Table of Contents)** - 展示整体结构
3. **标题 + 正文 (Title + Content)** - 最常用内容页
4. **左文右图 (Text Left + Image Right)** - 图文并茂
5. **右文左图 (Image Left + Text Right)** - 镜像布局
6. **双栏对比 (Two Column Compare)** - A/B 对比
7. **数据大字报 (KPI Big Number)** - 突出关键数字
8. **图片网格 (Image Grid)** - 展示多张图片
9. **引用金句 (Quote)** - 突出重要观点
10. **结束页 (Ending)** - 演示结尾

## 质量检查清单

- **P0 阻断级**：文件可打开、文字可读、封面结束页存在
- **P1 严重级**：字体统一、配色一致、布局不重复
- **P2 一般级**：文字量适中、图片清晰、页码连续
- **P3 锦上添花**：视觉冲击力、关键数据突出、行动号召

## 与现有工具协作

- **本 Skill (ppt-skill)**：python-pptx 本地操作，适合精细化控制、编辑、模板套用
- **Skywork-ppt**：远程 API 生成，适合快速从主题生成完整 PPT

**调度建议**：
- 需要从零生成内容丰富的 PPT → 优先用 `ppt-skill` 工作流 A
- 编辑已有 .pptx 文件 → 必须用 `ppt-skill` 工作流 B
- 需要 AI 自动配图 + 生成 → 可先用 `ppt-skill` 生成内容，再按需调用其他图片生成能力

## 依赖安装

Skill 脚本需要以下 Python 包：
```bash
pip install python-pptx
```

Agent 会在首次使用时自动安装。

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request。主要改进方向：
- 增加更多布局类型
- 优化截图适配工作流
- 扩展主题色包（但仍保持预设限制）
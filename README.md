# pptx-agent-toolkit

**Agent-Native PPTX Engineering Library** — Built for AI agent pipelines, not template filling.

22 unified layouts | 4 brand palettes | WCAG-AA color safety | Native CJK font support

---

## Why This Exists

Existing PPTX tools fall into two camps:

- **Template fillers** (Gamma, Beautiful.ai): Fast but shallow. They swap text into fixed designs. No understanding of your content, no quality guarantees, poor CJK support.
- **Raw python-pptx**: Full control but tedious. Every shape, color, and alignment is manual.

**This library is the third way**: a **pure-function API** designed for AI agents to programmatically construct, analyze, and enrich presentations with automated quality validation. It's the toolkit that powered 20+ production presentations at Shanghai Jiao Tong University.

## Quick Start

```bash
pip install pptx-agent-toolkit
```

```python
from pptx_agent_toolkit.builder import init_prs, new_slide, add_tb, set_p, add_rect, PALETTES
from pptx_agent_toolkit.color_utils import palette_aa_check

# Validate your colors pass WCAG AA
warnings = palette_aa_check(PALETTES["premium"])
assert len(warnings) == 0, f"Color issues: {warnings}"

# Build a slide
prs, C, font = init_prs({"palette": "premium"})
slide = new_slide(prs, C)
add_rect(slide, 0, 0.6, 13.333, 5.0, C["primary"])
tb = add_tb(slide, 1.2, 1.6, 10.0, 1.2)
set_p(tb.text_frame.paragraphs[0], "Hello, Agent!", C["white"], 40, bold=True)

prs.save("output.pptx")
```

See [`examples/demo_basic.py`](examples/demo_basic.py) for a complete 5-slide deck.

## What's Inside

### Builder — Core construction API
- **Pure functions**, zero classes. Each call is stateless and composable.
- 4 brand palettes: `premium`, `modern`, `dark`, `warm`
- Slide helpers: `new_slide`, `add_tb`, `set_p`, `add_p`, `add_rect`, `add_rrect`, `add_oval`, `place_image`, `add_footer`
- 16:9 widescreen (13.333"x7.5") by default

### Layout Enricher — Analyze & improve existing decks
- **22 unified layouts** merged from 3 sources (cookbook + layout_engine + design_mcp), deduplicated
- `EnrichmentEngine`: scores slide richness (0-100), recommends better layouts, detects image needs
- CLI: `python -m pptx_agent_toolkit.layout_enricher analyze input.pptx`

### Color QA — WCAG-AA contrast auditing
- `ColorQA`: audits entire PPTX files for text/background contrast
- `palette_aa_check`: validates color pairs before generation
- `find_contrast_fix`: auto-suggests safe color alternatives
- Built on WCAG 2.1 relative luminance math (zero external dependencies beyond stdlib)

### Image Tools
- Extract embedded images from PPTX, DOCX, PDF sources
- Batch image insertion with caption support

### Layout Engine — Alternative construction API
- `LayoutEngine` class with template-based slide construction
- Palette validation with auto-fix

## Layout Catalog (22 layouts)

| Category | Layouts | Visual Weight |
|----------|---------|--------------|
| **Structural** | cover, section-header, closing | 4-5 |
| **Content** | content, cards, side-by-side, comparison, bullets, numbered | 1-3 |
| **Visual** | hero, gallery, process, timeline, key-takeaways, quote, stats | 3-5 |
| **Hybrid** | text-with-image, image-with-text, split, carousel, table | 2-4 |
| **Special** | cta, hook, single-point | 3-5 |

## Design Principles

1. **Agent-first**: every function is stateless, serializable (JSON configs), and safe for automated construction
2. **No class wrappers**: flat functions avoid the XML corruption bugs common in OOP wrappers around python-pptx
3. **CJK native**: all fonts set across all 4 OOXML font slots (ascii, hAnsi, eastAsia, cs)
4. **Quality built-in**: WCAG-AA validation happens at generation time, not as an afterthought
5. **Composable**: each function does one thing. Combine them to build any layout.

## Requirements

- Python >= 3.10
- python-pptx >= 0.6.21
- Pillow >= 9.0

## Production Usage

This toolkit has been used to generate:
- Academic defense presentations (protein miniaturization, genetics experiments)
- Cultural education decks (Chinese cuisine, regional culture)
- Course summary presentations with 30+ slides
- Enriched/beautified versions of existing PPTX files

All generated presentations passed:
1. Programmatic WCAG-AA color audit
2. Independent visual review (cross-model validation)
3. Layout diversity constraints (no repeating layout >2 consecutive slides)

## License

MIT

---

# pptx-agent-toolkit（中文）

**面向 AI Agent 的原生 PPTX 工程库** — 为 Agent 流水线设计，而非模板填充。

22 种统一版式 | 4 套品牌调色板 | WCAG-AA 颜色安全 | 原生中文字体支持

## 为什么需要这个库

现有的 PPT 工具分成两类：

- **模板填充型**（Gamma、Beautiful.ai）：快但浅。把文字塞进固定设计，不理解你的内容，无质量保证，中文支持差。
- **裸 python-pptx**：完全可控但繁琐。每个形状、颜色、对齐都要手写。

**这个库是第三条路**：一套**纯函数式 API**，专为 AI Agent 编程式构造、分析、丰富演示文稿而设计，自带质量验证。已在上海交通大学产出 20+ 个高质量演示文稿。

## 快速开始

```bash
pip install pptx-agent-toolkit
```

```python
from pptx_agent_toolkit.builder import init_prs, new_slide, add_tb, set_p
from pptx_agent_toolkit.color_utils import palette_aa_check, PALETTES

# 验证调色板通过 WCAG AA 无障碍标准
warnings = palette_aa_check(PALETTES["premium"])
# 构建幻灯片
prs, C, font = init_prs({"palette": "premium"})
slide = new_slide(prs, C)
tb = add_tb(slide, 1.2, 1.6, 10.0, 1.2)
set_p(tb.text_frame.paragraphs[0], "你好，Agent！", C["primary"], 40, bold=True, font=font)
prs.save("output.pptx")
```

详见 [`examples/demo_basic.py`](examples/demo_basic.py)，一个完整的 5 页演示文稿示例。

## 核心模块

### Builder — 核心构建 API
- **纯函数**，零类封装。每次调用无状态、可组合。
- 4 套品牌调色板：`premium`（暖学术）、`modern`（现代蓝）、`dark`（深色科技）、`warm`（暖色文化）
- 16:9 宽屏（13.333"x7.5"）默认尺寸

### Layout Enricher — 分析并优化现有 PPT
- **22 种统一版式**，从 3 个来源合并去重（cookbook + layout_engine + design_mcp）
- `EnrichmentEngine`：评分（0-100）、推荐版式、检测图像需求
- CLI：`python -m pptx_agent_toolkit.layout_enricher analyze input.pptx`

### Color QA — WCAG-AA 颜色无障碍审查
- `ColorQA`：审计整个 PPTX 文件的文字/背景对比度
- `palette_aa_check`：生成前验证颜色配对
- `find_contrast_fix`：自动推荐安全的替代颜色

### Image Tools — 图像提取与插入
- 从 PPTX、DOCX、PDF 中提取嵌入图像
- 批量插入图像，支持题注

### Layout Engine — 备选构建 API
- `LayoutEngine` 类，基于模板的幻灯片构建
- 调色板验证与自动修复

## 版式目录（22 种）

| 分类 | 版式 | 视觉权重 |
|------|------|---------|
| **结构性** | 封面、章节分隔、结束页 | 4-5 |
| **内容型** | 标准内容、卡片、并排、对比、要点、编号 | 1-3 |
| **视觉型** | 大图、图库、流程、时间线、关键要点、引用、数据 | 3-5 |
| **混合型** | 左文右图、左图右文、左右分割、轮播、表格 | 2-4 |
| **特殊** | CTA、钩子、单点 | 3-5 |

## 设计原则

1. **Agent 优先**：每个函数无状态、可序列化（JSON 配置），适合自动化构建
2. **零类封装**：纯函数避免 OOP 包装 python-pptx 时常见的 XML 损坏 bug
3. **中文本地化**：所有字体在 4 个 OOXML 字体槽（ascii、hAnsi、eastAsia、cs）同时设置
4. **质量内建**：WCAG-AA 验证在生成时进行，而非事后补救
5. **可组合**：每个函数只做一件事，组合起来构建任意版式

## 生产使用记录

本工具包已用于生成：
- 学术答辩演示文稿（蛋白质微型化、遗传学实验）
- 文化教育课件（中国地方菜系、地域文化）
- 30+ 页课程总结报告
- 已有 PPTX 的美化增强版本

所有生成的演示文稿均通过：
1. 程序化 WCAG-AA 颜色审计
2. 独立视觉审查（跨模型交叉验证）
3. 版式多样性约束（相同版式不可连续超过 2 页）

## 许可

MIT License

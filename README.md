# vsdx-replicate

**位图流程图 → 可编辑 Visio (.vsdx) 的 1:1 复刻技能**，面向 ZCode / Claude Code 类 Agent 环境。

给定一张位图截图或矢量导出图，按以下管线产出外观 1:1、全部形状独立可选中可编辑的 `.vsdx`：

```
像素级提取(调色板→色块bbox→线网→放大目检) → 位图素材裁剪
→ Visio COM 渲染(几何+纯文本) → XML 后处理(富文本/字体/粗斜体/下标)
→ PDF 栅格化分区对比 → 迭代修正 → 交付
```

核心价值不在"生成 vsdx"，而在沉淀了这一流程里所有的坑：`pywin32` 驱动 Visio COM 的坐标/单位约定、多 run 富文本必须走 OPC XML 后处理（COM 的 `CharProps` 浮点值会被强转整型、颜色 >32767 溢出 I2）、虚线/跳线(hop)/箭头起止的线网检测方法、禁用 `page.Export` 的对比链路（Visio 会拉伸到缓存画布）等。

## 目录结构

```
vsdx-replicate/
├── SKILL.md                # 完整方法论与全部踩坑记录（Agent 技能入口）
└── scripts/                # 参考实现（对一张真实论文流程图跑通）
    ├── extract.py          # Stage 1：调色板 + 色块 bbox 提取（近邻色并查集合并）
    ├── trace2.py           # 线网追踪：实线/虚线/箭头/跳线(hop)
    ├── render.py           # Stage 3：COM 渲染，几何 + 纯文本 + runs dump
    ├── patch_text.py       # Stage 4：vsdx OPC XML 后处理富文本
    └── export_compare.py   # Stage 5：PDF 导出 → 栅格化 → 分区差异对比
```

脚本中的 `SRC`（源图路径）、输出文件名 `out/output.vsdx`、页面宽高 `W/H` 为占位值，替换为目标图即可起步。

## 环境要求

- Windows + 本机安装 Microsoft Visio（COM 自动化是渲染核心）
- Python 3.x，依赖：

```bash
pip install pywin32 opencv-python pillow pymupdf
```

无 Visio 的机器可改走手写 OPC XML 路线（参考 Microsoft [MS-VSDX] 规范），本技能不适用批量转换场景（请用 drawio CLI `-f vsdx`）。

## 验收标准（详见 SKILL.md）

- Visio 打开无修复提示，全部形状独立可选中/可编辑
- 分区对比：位图区 <3%，线网/容器错位 ≤2px
- 文字逐字一致，粗斜体/下标/竖排/灰字齐全，无换行溢出

## License

MIT

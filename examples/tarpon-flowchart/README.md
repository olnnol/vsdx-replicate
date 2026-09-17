# 示例：TARPON 流程图画（论文配图 → 可编辑 Visio）

一个完整的 worked example —— 把一篇开放获取论文里的流程图（位图）复刻为 19 个形状、全部可编辑的 Visio 文件。

| 文件 | 说明 |
|---|---|
| `source.png` | 位图输入（论文 Fig 1a，2531×1316） |
| `replica.vsdx` | 复刻产物（Visio 直接打开，所有形状独立可选中/编辑） |
| `render_example.py` | 该图的场景数据与渲染脚本（Stage 3，1 px = 0.1 mm） |
| `out/` | 运行产物（已 gitignore） |

效果对比见仓库根目录 [docs/](../../docs/) 下的 `demo-before-after.png` 与 `demo-details.png`。

## 复跑

需要 Windows + 已安装 Visio + `pip install pywin32 opencv-python pillow pymupdf`。

```bash
cd examples/tarpon-flowchart

# Pass 1: 几何 + 纯文本 → out/replica.vsdx + out/text_runs.json
python render_example.py

# Pass 2: OPC XML 富文本（字体/字号/颜色）
python ../../scripts/patch_text.py out/replica.vsdx out/text_runs.json

# 可选：导出 PDF 栅格化，与原图做差异对比
python ../../scripts/export_compare.py --src source.png --vsdx out/replica.vsdx
```

## 这张图里值得看的复刻点

- **形状体系**：实心圆角矩形（r=60px，无描边）、扁平菱形（判断节点）、实心三角楔形连接件、凹底隐形箭头；
- **曲线箭头**：三点拟合圆（圆心/半径实测）后按 48 段折线绘制 —— Visio `DrawArcByThreePoints` 的取弧语义与直觉不符，折线法可控且平滑；
- **字体标定**：原图为 Montserrat 风格几何无衬线体，实测字符宽度比与笔画宽度（5px stem）后选定 Arial（宽度比 1.03、stem 一致、双层 a 结构一致），字号按帽高反推为 60px；
- **文字定位**：双行文字行距 73px（1.22 倍字号），蓝菱形文字整体低于几何中心 13px —— 通过 Top/Bottom 边距不对称实现像素级对位；
- **调色板**：`084D41` 深绿 / `D81A61` 品红 / `3A82C4` 蓝 / `FEC110` 金黄 / `231F20` 近黑，全部按源图取样。

## 复刻指标

- 分区像素对比：整体 mean abs diff 6.6，`px>60` 占比 3.4%（差异主要来自字体渲染引擎）；
- 文字行位、行距与原图误差 ≤3px；曲线箭头轨迹误差 ≤5px；
- Visio 打开无修复提示，19 个形状全部独立可选中/可编辑。

## 图片来源与许可

> Deimler N, Ho DV, Paul N, Gill Z, Baumann P. **TARPON—A Telomere Analysis and Research Pipeline Optimized for Nanopore sequencing.** *PLOS Computational Biology*, 2026. DOI: [10.1371/journal.pcbi.1013915](https://doi.org/10.1371/journal.pcbi.1013915)
>
> Figure used: Fig 1(a). Licensed under [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/) — reproduced with attribution to the original authors.
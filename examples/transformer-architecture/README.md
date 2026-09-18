# 示例：Transformer 架构图（论文经典结构 → 可编辑 Visio）

一个完整的 worked example —— 把 "Attention Is All You Need" (Vaswani et al., 2017) 的
Transformer 编码器-解码器架构图（现代重绘版）复刻为 **140 个形状、全部可编辑** 的 Visio 文件。

| 文件 | 说明 |
|---|---|
| `source.png` | 位图输入（架构图，1426×1500） |
| `replica.vsdx` | 复刻产物（Visio 直接打开，所有形状独立可选中/编辑） |
| `render_example.py` | 该图的场景数据与渲染脚本（Stage 3，1 px = 0.1 mm） |
| `out/` | 运行产物（已 gitignore） |

效果对比见仓库根目录 [docs/](../../docs/) 下的 `demo-before-after.png` 与 `demo-details.png`。

## 复跑

需要 Windows + 已安装 Visio + `pip install pywin32 opencv-python pillow pymupdf`。

```bash
cd examples/transformer-architecture

python render_example.py                                            # 几何 + 纯文本
python ../../scripts/patch_text.py out/replica.vsdx out/text_runs.json   # 富文本
python ../../scripts/export_compare.py --src source.png --vsdx out/replica.vsdx
```

## 这张图里值得看的复刻点

- **双层容器**：编码器（橙）与解码器（绿）大圆角矩形，内部再嵌套虚线残差框（5 个，圆角 40px、2px 虚线）；
- **V/K/Q 分配器**：主轴在分支杆上的圆角弯头 + 三个箭头，Q 支路由 Norm 顶部 S 弯引入，V/K 由编码器输出横线喂入；
- **编码器输出路由**：上出容器 → 右转 → 下行 → 箭头进入解码器虚线框边界（终止箭头）→ 框内馈线，四段路径与原图一致；
- **图形符号**：⊕ 加法圆（十字臂 0.31r）、位置编码圆（正弦波贯穿圆面）；
- **字体**：原图为 draw.io 默认的 Verdana Bold——经字形宽度比对（12 字形中 11 个完全吻合）确认，复刻直接使用系统自带 Verdana Bold，字号按帽高反推；
- **配色**：draw.io 默认色板 `D5E8D4/82B366`、`FFF2CC/D6B656`、`DAE8FD/6C8EBF`、`F8CECC/B85450`、`D79B00`。

## 复刻指标

- 整体像素差 mean 8.4，`px>60` 占比 4.4%（残余为字体渲染引擎差异）；
- 文字字形级对位：逐字符位置误差 1-2px（"Norm"/"Multi-Headed"/"Embeddings/" 等逐字形核验）；
- Visio 打开无修复提示，140 个形状全部独立可选中/可编辑。

## 图片来源与许可

> 架构图源文件：dvgodoy, *"Transformer, full architecture"* 重绘版，
> [Wikimedia Commons](https://commons.wikimedia.org/wiki/File:Transformer,_full_architecture.png)，[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)。
> 原始结构出自 Vaswani A, et al. *"Attention Is All You Need"* (arXiv:1706.03762, NeurIPS 2017)。
> 本示例图片按 CC BY 4.0 使用并对原作者署名；本目录产物（replica.vsdx 等）仅用于演示本技能。
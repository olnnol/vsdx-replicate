---
name: vsdx-replicate
description: >
  把流程图/架构图/论文模块图的图片 1:1 复刻为可编辑的 Microsoft Visio .vsdx 文件
  （原生形状+文字+连接线，不是贴图）。通过像素级提取（调色板聚类、线网追踪、
  箭头/文字清单）+ Visio COM 自动化 + VSDX XML 富文本后处理 + PDF 栅格化对比
  闭环实现。Use when user asks to "复刻流程图", "重绘为 Visio", "转成 vsdx",
  "1:1 重绘", "visio 复刻", or mentions "vsdx-replicate".
---

# VSDX Replicate：图片 → 可编辑 Visio 流程图 1:1 复刻

参考实现（已跑通、可直接改参复用）：本目录 `scripts/`，原始工作副本在
参考实现（已跑通、可直接改参复用）：本目录 `scripts/`。脚本中的 `SRC` 路径、输出文件名
`output.vsdx` 与页面宽高 `W/H` 为占位值，替换为目标源图即可起步。

## 适用与不适用

- 适用：Windows + 本机装有 Visio（COM 可用），源图为位图截图或矢量导出图，
  要求产出可编辑 .vsdx 且外观 1:1。
- 不适用：批量转换（用 drawio CLI `-f vsdx`）、无 Visio 的机器
  （改走手写 OPC XML，参考 Microsoft [MS-VSDX] 规范）。

## 核心认识

**瓶颈不在"生成 vsdx"，而在"从源图精确提取每个元素的位置/尺寸/颜色/文字/箭头"。**
生成端（COM + XML）是一次性搭好的管线；提取端每张图都要做。顺序：

```
提取(调色板→色块bbox→线网→放大目检) → 裁位图素材 → COM渲染(几何+纯文本)
→ XML后处理(富文本) → PDF栅格化对比 → 迭代修正 → 交付
```

## Stage 0 环境检查

```python
# 需全部通过；pywin32 驱动 Visio COM 是渲染核心
import win32com.client, cv2, fitz  # pip: pywin32 opencv-python pillow pymupdf
app = win32com.client.Dispatch('Visio.InvisibleApp')  # 打印 app.Version 后 Quit
```
卡死/弹窗处理：`app.AlertResponse = 2` 压制对话框；异常僵死用
`taskkill //IM VISIO.EXE //F` 清进程后重试。

## Stage 1 像素级提取（每张图必做）

1. **读图**：`PIL.Image.open`（cv2.imread 在 Windows 中文路径会失败）。
   记录 W×H，全流程统一用源图像素坐标。
2. **调色板**（`scripts/extract.py` 思路）：精确颜色直方图（≥2000px）+
   并查集合并距离≤4 的近邻色（源图有压缩噪点，同一填充会散布在多个精确值上，
   逐值统计会切碎调色板——这是最大的坑）。对每个色簇跑连通域得色块 bbox。
3. **线网追踪**（`scripts/trace2.py` 思路）：
   - 黑色掩码（g<90）排除位图区域后，水平/垂直形态学开运算（61px 核）提线段，
     共线合并；合并线内的 4~45px 空洞 = 跳线（hop）位置。
   - 虚线：灰掩码要**含 g<110 的深灰核**（虚线核心常是 #5A5A5A 左右，
     只取 110-215 会漏核、只取到抗锯齿晕）。
   - 容器边框会被当线段检出，按已知 bbox 过滤。
4. **目检定歧**：对所有容器/交汇区生成裁剪图（关键区 2~4x 放大）逐一 Read 目检，
   确认：箭头方向与起止点、T 型分叉 vs 跨线（有无 hop）、虚线网络的源与汇、
   圆圈符号（⊕⊖~）内部笔画、竖排文字读向、括号/小字。**不要靠猜**，
   本方法的可靠性来自"线网数据 + 人眼/视觉核验"双通道。
5. **颜色定点**：边框/描边用"穿过边缘的扫描线"取核色（单点探针易落在填充上）；
   文字灰度用分位 darkest 提取。

## Stage 2 位图素材裁剪

照片级内容（实验图、热力图、装饰性素描）按 bbox 裁剪嵌入，背景色与容器填充一致
可无缝拼接。**边距要给足（≥8px）**：先用"笔画掩码逐行统计"确认真实外沿
（形态学闭运算估计的 bbox 会偏小，曾把小球底部切掉 20px）。
裁剪下界注意不要吃进相邻文字的抗锯齿。

## Stage 3 COM 渲染（`scripts/render.py` = 场景数据 + 渲染器）

坐标约定：**1 px = 0.1 mm**，页面 `fx(W) × fx(H)`（W,H=源图宽高，直接用 mm 公式）。
所有元素按 z 序绘制：容器 → 内框 → 位图 → 线网 → 圆圈符号 → 文本。

**COM API 的坑（全部踩过）：**

| 坑 | 正确做法 |
|---|---|
| `DrawRectangle/DrawLine/DrawOval/DrawArcByThreePoints` 不收公式字符串 | 传**英寸浮点**：`px/254.0`、`(H-py)/254.0`（Visio Y 轴向上） |
| 方法名 | `DrawArcByThreePoints`（不是 By3Points） |
| PageHeight | 直接 `fx(H)`；误用翻转公式会得到 0 |
| 文本块默认=形状尺寸 | 隐形文本矩形必须显式设 `TxtWidth/TxtHeight`，否则逐字换行成竖列 |
| 旋转文本 | `TxtAngle='90 deg'`=逆时针（读向自下而上）；旋转时 **TxtWidth/TxtHeight 互换** |
| `Characters.CharPropsRow` 恒返回行 0 | 逐区间格式化全部互相覆盖（加粗丢失/字号被最后一个区间统一） |
| `CharProps` 索引属性 put | Python 语法不支持，用 `obj._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, True, idx, val)`；但**浮点值会被强转整型**（字号变 1pt）、**颜色 >32767 溢出 I2** —— 所以富文本别走 COM，走 Stage 4 XML |
| `FONTTOID("宋体")` | 中文 font 名返回 #VALUE!；用 `doc.Fonts('SimSun').ID`（=5，Times New Roman=23） |

因此 render.py 只做：几何 + 纯文本 + 段落对齐 + Txt 块尺寸，形状命名
（容器/框给语义名，文本形状自动编号 TEXT01..），并把 runs 结构 dump 成
`text_runs.json`。

## Stage 4 XML 后处理（`scripts/patch_text.py`，本管线最关键的经验）

vsdx = OPC zip；改 `visio/pages/page1.xml` 后重打包。对每个命名形状：

1. 定位形状：`<Shape ID='n' ... Name='TEXT01'` —— **只改 .Name 时 NameU 保留旧名，
   必须匹配 `Name=`**（注意别误匹配 NameU=）。
2. 生成 `<Section N='Character'>` + `<Text>`（`<cp IX='k'/>` 切换 run）。
   **Row 必须写完整单元格集**，缺省值会被解释为 0%，导致该 run 文字零宽叠堆成乱码：

```xml
<Row IX='0'>
 <Cell N='Font' V='SimSun'/>            <!-- 或 Times New Roman，直接写名字 -->
 <Cell N='Color' V='#7F7F7F'/>          <!-- 省略则继承主题黑 -->
 <Cell N='Style' V='1'/>                <!-- 位: 1粗 2斜 3粗斜 -->
 <Cell N='Case' V='0'/>
 <Cell N='Pos' V='2'/>                  <!-- 2=下标（Visio 会自动再缩小约0.7，字号先×1.25 补偿） -->
 <Cell N='FontScale' V='1'/>
 <Cell N='Size' V='0.1732283464566929'/> <!-- 英寸 = px/254 -->
 <Cell N='DblUnderline' V='0'/>... <Cell N='ColorTrans' V='0'/>
 <Cell N='AsianFont' V='Themed' F='THEMEVAL()'/>
 <Cell N='ComplexScriptFont' V='Themed' F='THEMEVAL()'/>
</Row>
```

3. 形状若没有 Character 节（COM 未写过格式），**在 `<Text>` 前插入**——
   `re.sub` 找不到目标会静默不替换，务必走 else 分支。
4. Text 内换行就是字面 `\n`；`&<>` 需转义；行内相邻同格式 run 合并。

## Stage 5 导出对比闭环（`scripts/export_compare.py`）

**禁止 `page.Export(png/emf)`**：Visio 会把页面拉伸到上次导出会话缓存的画布
（比例错乱且跨会话持久）。正确链路：

```
doc.ExportAsFixedFormat(1, 'out.pdf', 1, 0)   # 页面几何精确
→ PyMuPDF get_pixmap(zoom=2800/pdf宽pt) → 2800px PNG
→ 与原图分区对比：mean abs diff + 每区 >60 像素占比 + 三联图
  （上原图/下渲染 或 左右并排），逐区 Read 目检 → 修 render.py 场景数据 → 重跑
```

注意：源图自带压缩噪点，纯白区也有 1~3 级抖动，整体 mean diff 永远不会到 0；
以"分区占比 + 目检"为准。位图嵌入区应 <3%，文字区 10~25% 属正常
（不同排版引擎的宋体渲染差异，无法逐像素相同）。


## 图标与小装饰：位图会糊，重画为矢量

源图里的小图标（扫帚、时钟、网格、柱状图、热力网格、仪表盘、盾牌等）若按 bbox
裁剪为位图嵌入，在 Visio 缩放/导出时会发虚。正确做法是**矢量重画**：

- 几何类图标全部用原生形状重画：`DrawOval`（表盘/节点）、`DrawPolyline`
  （闭合多边形可填充：等距立方体、盾牌、警告灯圆顶半椭圆）、`DrawRectangle`
  （柱状图、热力网格逐格填色——颜色直接从源图逐格采样）、`DrawArcByThreePoints`
  （仪表盘弧线）、贝塞尔/多段折线近似曲线（钟形曲线逐点算 gaussian）。
- **点划曲线**（虚线轨迹）：不要用 `LinePattern=2` 的长虚线（折线段上会变成
  折线感大虚线），改为沿路径取点后**隔点连线**（短划 + 空隙 + round cap）。
- **小箭头**：Visio `EndArrow` 头部随线宽、下限偏大，小图标里的箭头要
  手绘小三角形（`shape` 填充三点）。
- 数据类图像（实验图、热力图照片）保持位图裁剪——那是照片级内容，矢量无意义。
- 图标绘制封装成独立 icons.py 模块（Ctx 类封装 page + 局部坐标偏移），
  render.py 里用 `icons.X(icons.Ctx(page, H, x0, y0))` 替换 `image()` 调用。


## 矢量输出链路（SVG / EMF / PDF）

设计类新图（非复刻）同一管线可产出全矢量交付物：

1. **PDF（基准，精确）**：`doc.ExportAsFixedFormat(1, out.pdf, 1, 0)`，页面几何与
   VSDX 完全一致。
2. **SVG**：PyMuPDF `page.get_svg_image(text_as_path=True)`——文字转矢量轮廓，
   自包含。**坑：根元素 width/height 是无单位数字（=pt 值）**，其他工具会按 96dpi
   误读导致缩放/裁切；导入 LO 前改为显式 `width="152.4mm" height="111mm"`。
3. **EMF**：
   - Visio `page.Export('*.emf')` 尺寸元数据不可信（跨会话画布缓存，实测
     109×76mm/152×111 目标），放弃。
   - LibreOffice `soffice --headless --convert-to emf x.svg` 内容会被 Draw 页面
     裁切/缩放（141.7×100.3mm 且轻微纵横比失真），仅作备选。
   - **首选 PowerPoint COM**：`Presentations.Add(False)` → **先 `Slides.Add(1,12)`
     （空演示文稿 0 页，直接 Slides(1) 会报错）** → `PageSetup.SlideWidth/Height`
     设为 PDF 的 pt 尺寸 → `Shapes.AddPicture(svg, ...)` 铺满 → `Slide.Export
     (emf, 'EMF')`。实测 frame 152.47×110.84mm（目标 152.4×111），内容零裁切。
4. PNG 预览：PyMuPDF 按 pt 宽度比例栅格化。

## 适配新图 checklist

1. 全局替换 `SRC` 路径（extract/trace2/export_compare）与页面 W/H（render.py）。
2. 跑 Stage 1，把输出的容器/内框 bbox 与实测颜色填进 render.py 场景区。
3. 位图素材：改裁剪清单 + `image()` 调用。
4. 线网：对照 trace 输出 + 目检裁剪图重写 `wire()/hop()` 清单。
5. 文本：每个文本块给 bbox（宁宽勿窄，窄必换行）+ runs
   （`S(文字, 字号px, f='S'|'T', b, i, sub, col)`）。
6. 迭代 2~4 轮到分区目检通过，交付 vsdx + side_by_side 对比图。

## 验收标准

- Visio 打开无修复提示，105± 形状全部独立可选中/可编辑；
- 分区对比：位图区 <3%，线网/容器无可见错位（≤2px）；
- 文字内容逐字一致，加粗/斜体/下标/竖排/灰字齐全，无换行溢出；
- 箭头起止、hop 位置、虚实线分类与原图一一对应。

# rf-residual-kriging-suite

## 关于 / About

**中文**

这是一个面向土壤三普成果修改场景的 Trae skill 归档仓库，核心功能是把 **随机森林预测栅格 + 残差普通克里金校正**、下限校正、分级重分类、精度重算和统计表导出串成一套可复用流程。

**English**

This repository archives a Trae skill for soil-survey revision workflows. Its core purpose is to turn **RF prediction raster + residual ordinary kriging correction**, optional lower-bound correction, grading reclassification, accuracy recalculation, and summary-table export into a reusable pipeline.

## 中文说明

这个 skill 不是单纯的算法示例，而是用于土壤三普成果修改的完整工作流，主要包含：

- 对 RF 预测栅格做残差普通克里金叠加校正
- 可选的下限校正
- 校正前后精度重算
- 按分级标准重分类为 1~5 级分级栅格
- 样点分级统计、制图面积分级统计
- 验证集点表、训练集点表导出
- 土地利用、乡镇、地类分级面积统计补充表

### 适用场景

- 已经有 RF 预测栅格
- 已经有验证集点表和残差表
- 需要在 ArcGIS Pro / arcpy 环境里完成残差克里金校正
- 需要重新输出验证精度、分级结果和统计表
- 需要按公报面积统一口径做等比例缩放

### 主要输入

- RF 预测栅格：`属性_预测结果.tif`
- 验证集残差表：至少包含属性、经度、纬度、实际值、预测值、残差
- 样点表：至少包含 `longitude`、`latitude`、各属性字段、`土地利用`
- 分级标准表：`2属性分级标准表.xlsx`
- 土地利用与乡镇矢量：`DLMC`、`一级地`、`XZQMC`
- 公报面积：用于面积缩放和一致性校验

### 核心流程

1. 从验证集表中筛选出指定属性的验证点
2. 生成点要素并投影到 RF 栅格坐标系
3. 对残差进行普通克里金插值
   - 变差函数：球状函数
   - 搜索邻点数：12
   - 环境参数：`snapRaster`、`extent`、`cellSize` 与 RF 栅格一致
4. 叠加校正：`校正栅格 = RF 栅格 + 残差克里金栅格`
5. 必要时执行下限校正，避免结果低于设定阈值
6. 重新提取验证点结果，计算 `R²` 和 `RMSE`
7. 按分级标准重分类，生成 1~5 级分级栅格
8. 输出分级对比表、验证集点表、训练集点表与补充统计表

### 推荐输出目录

- `01_残差克里金栅格`
- `02_叠加校正栅格`
- `03_分级栅格`
- `04_表格`
- `05_验证集点表`
- `05_训练集点表`
- `06_补充统计表`

### 推荐脚本入口

- `rf_residual_kriging_pipeline.py`
- `export_corrected_validation_points.py`
- `export_training_points.py`
- `landuse_township_tables.py`

示例运行方式：

```powershell
& "C:\installsoft\gispro35\bin\Python\scripts\propy.bat" "G:\...\rf_residual_kriging_pipeline.py"
```

### 仓库内容

本仓库当前保存的是这个 Trae skill 的定义文件：

- `rf-residual-kriging-suite/SKILL.md`

---

## English

This repository archives a Trae skill for **RF prediction raster correction with residual ordinary kriging**.

The skill is designed to turn the full soil-survey correction workflow into a reusable pipeline, including:

- residual ordinary kriging correction on the RF prediction raster
- optional lower-bound correction
- recalculation of post-correction accuracy
- reclassification into 1-5 grade rasters using a grading standard table
- sample-based grading statistics and mapped-area statistics
- export of validation-point tables and training-point tables
- supplementary summary tables for land use, townships, and land-type grading areas

### When to use it

- You already have an RF prediction raster
- You already have validation-point residual tables
- You need residual kriging correction in ArcGIS Pro / arcpy
- You need updated accuracy metrics, grading outputs, and summary tables
- You need area scaling that matches published report totals

### Main inputs

- RF prediction raster: `属性_预测结果.tif`
- Validation residual table: at minimum, attribute, longitude, latitude, actual value, predicted value, and residual
- Sample-point table: at minimum, `longitude`, `latitude`, attribute fields, and `土地利用`
- Grading standard table: `2属性分级标准表.xlsx`
- Land-use and township vectors: fields such as `DLMC`, `一级地`, and `XZQMC`
- Published report area totals for scaling and consistency checks

### Core workflow

1. Filter validation points by attribute from the merged validation table
2. Generate point features and project them into the RF raster coordinate system
3. Perform ordinary kriging on residuals
   - Variogram model: spherical
   - Search neighbors: 12
   - Environment settings: `snapRaster`, `extent`, and `cellSize` aligned with the RF raster
4. Apply correction: `corrected raster = RF raster + residual kriging raster`
5. Apply lower-bound correction when needed
6. Extract validation points again and compute `R²` and `RMSE`
7. Reclassify the corrected raster into grades 1-5
8. Export comparison tables, validation/training point tables, and supplemental statistics

### Recommended output folders

- `01_residual_kriging_raster`
- `02_corrected_raster`
- `03_grading_raster`
- `04_tables`
- `05_validation_points`
- `05_training_points`
- `06_supplementary_tables`

### Suggested script entry points

- `rf_residual_kriging_pipeline.py`
- `export_corrected_validation_points.py`
- `export_training_points.py`
- `landuse_township_tables.py`

Example:

```powershell
& "C:\installsoft\gispro35\bin\Python\scripts\propy.bat" "G:\...\rf_residual_kriging_pipeline.py"
```

### Repository contents

This repository currently stores the Trae skill definition file:

- `rf-residual-kriging-suite/SKILL.md`
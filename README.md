# RF 残差克里金校正与统计全流程

这是一套完整的"随机森林预测栅格 + 残差克里金插值校正"工具链，用于土壤属性制图。

## 功能特性

- ✅ **残差克里金插值**：对验证集残差进行普通克里金插值
- ✅ **叠加校正**：RF 栅格 + 残差克里金栅格 = 最终校正栅格
- ✅ **下限校正**：自动处理负值并应用制图下限
- ✅ **精度对比**：验证 RF 和校正后栅格的 R²、RMSE
- ✅ **分级重分类**：按土壤三普分级标准重分类
- ✅ **统计对比表**：样点统计 vs 制图统计对比
- ✅ **验证/训练集点表导出**：导出校正后的点表
- ✅ **补充统计表**：表1（土地利用对比）、表2（乡镇分级）、表3（地类分级）

## 脚本说明

### 统一入口（推荐）

```powershell
& "C:\installsoft\gispro35\bin\Python\scripts\propy.bat" "scripts/run_full_pipeline.py"
```

特点：
- 运行前显示完整配置信息
- 交互式确认
- 自动按顺序执行全流程
- 运行状态和结果总结

### 独立脚本

| 脚本 | 功能 |
|------|------|
| `rf_residual_kriging_pipeline.py` | 残差克里金插值 + 叠加校正 + 精度/分级对比 |
| `export_corrected_validation_points.py` | 导出校正后的验证集点表 |
| `export_training_points.py` | 导出校正后的训练集点表 |
| `landuse_township_tables.py` | 生成补充统计表（表1、表2、表3） |

## 输入数据要求

1. **RF 预测栅格**：`属性_预测结果.tif`
2. **验证集表**：包含属性、实际值、预测值、残差、经纬度
3. **样点表**：包含 longitude、latitude 和各属性值
4. **属性分级标准表**：土壤三普分级标准
5. **土地利用矢量**：`DLMC` 和 `一级地` 字段
6. **乡镇界矢量**：`XZQMC` 字段

## 输出结构

```
输出目录/
├── 01_残差克里金栅格/
├── 02_叠加校正栅格/
├── 03_分级栅格/
├── 04_表格/
│   ├── RF与叠加残差克里金_精度对比.xlsx
│   └── 叠加残差克里金_分级对比表.xlsx
├── 05_验证集点表/
├── 05_训练集点表_下限校正/
└── 06_补充统计表/
    ├── 表1_土地利用样点与制图对比.xlsx
    ├── 表2_耕地_乡镇分级统计.xlsx
    └── 表3_各地类分级统计.xlsx
```

## 环境要求

- **ArcGIS Pro**（含 Spatial Analyst 扩展）
- **ArcGIS Pro Python 环境**（包含 arcpy）
- **Python**：3.x（随 ArcGIS Pro 安装）

## 运行方式

### 使用统一入口

1. 配置 `scripts/run_full_pipeline.py` 中的数据路径
2. 运行：
```powershell
& "C:\installsoft\gispro35\bin\Python\scripts\propy.bat" "scripts/run_full_pipeline.py"
```

### 单独运行

```powershell
& "C:\installsoft\gispro35\bin\Python\scripts\propy.bat" "scripts/rf_residual_kriging_pipeline.py"
```

## 项目结构

```
rf-residual-kriging-suite/
├── SKILL.md              # Trae 技能描述
├── README.md             # GitHub 项目说明（本文件）
└── scripts/              # 核心脚本目录
    ├── run_full_pipeline.py
    ├── rf_residual_kriging_pipeline.py
    ├── export_corrected_validation_points.py
    ├── export_training_points.py
    └── landuse_township_tables.py
```

## 许可证

MIT License

## 联系方式

如遇问题，请查看 SKILL.md 获取详细使用说明。

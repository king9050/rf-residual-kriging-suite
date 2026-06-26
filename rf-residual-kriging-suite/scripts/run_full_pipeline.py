#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
统一的 RF 残差克里金校正与统计全流程运行脚本
包含运行前确认和引导
"""

import os
import sys
import subprocess
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def print_separator():
    print("=" * 80)


def print_info(title, content):
    print(f"\n📋 {title}")
    print("-" * 40)
    if isinstance(content, list):
        for item in content:
            print(f"  • {item}")
    else:
        print(f"  {content}")


def check_file_exists(path, description):
    if not path:
        return f"⚠️  {description} 未设置"
    if os.path.exists(path):
        return f"✅ {description}: {path}"
    else:
        return f"❌ {description}（文件不存在）: {path}"


def check_dir_exists(path, description):
    if not path:
        return f"⚠️  {description} 未设置"
    if os.path.isdir(path):
        files = [f for f in os.listdir(path) if f.lower().endswith('.tif')]
        return f"✅ {description}（{len(files)} 个栅格）: {path}"
    else:
        return f"❌ {description}（目录不存在）: {path}"


def build_default_config():
    """返回当前项目默认配置。后续迁移到配置文件时只需替换这里。"""
    gispro_python = r"C:\installsoft\gispro35\bin\Python\envs\arcgispro-py3"
    gispro_propy = r"C:\installsoft\gispro35\bin\Python\scripts\propy.bat"
    base_dir = r"c:\Users\HiWin10\Desktop\浮梁自验收\报告及数据修改skill\浮梁属性栅格克里金插值_自验收属性修改20260623"

    return {
        "gispro_python": gispro_python,
        "gispro_propy": gispro_propy,
        "base_dir": base_dir,
        "sample_xlsx": os.path.join(base_dir, "浮梁土壤属性表层样修正.xlsx"),
        "validation_xlsx": os.path.join(base_dir, "浮梁_验证集-随机森林最优方案样本验证集.xlsx"),
        "standard_xlsx": os.path.join(base_dir, "2属性分级标准表.xlsx"),
        "landuse_shp": os.path.join(base_dir, r"fl_landuse\浮梁土地利用.shp"),
        "township_shp": os.path.join(base_dir, r"fl_乡镇界\fl乡镇界.shp"),
        "rf_dir": os.path.join(base_dir, "预测属性栅格"),
    }


def collect_config_status(config):
    """收集输入资料状态；只做存在性检查，不执行 GIS 处理。"""
    checks = [
        ("file", config["sample_xlsx"], "样点表"),
        ("file", config["validation_xlsx"], "验证集表"),
        ("file", config["standard_xlsx"], "属性分级标准表"),
        ("file", config["landuse_shp"], "土地利用矢量"),
        ("file", config["township_shp"], "乡镇界矢量"),
        ("dir", config["rf_dir"], "预测属性栅格目录"),
    ]
    missing = []
    messages = []
    for kind, path, description in checks:
        exists = os.path.isfile(path) if kind == "file" else os.path.isdir(path)
        if not exists:
            missing.append(description)
        if kind == "file":
            messages.append(check_file_exists(path, description))
        else:
            messages.append(check_dir_exists(path, description))
    return messages, missing


def confirm_config(check_only=False):
    """显示配置信息并获取用户确认"""
    print_separator()
    print("🚀 RF 残差克里金校正与统计全流程 - 运行配置")
    print_separator()

    config = build_default_config()
    gispro_python = config["gispro_python"]
    gispro_propy = config["gispro_propy"]
    
    print_info("🔧 GISPRO 环境", [
        f"Python 环境目录: {gispro_python}",
        f"ProPy 脚本路径: {gispro_propy}",
        f"环境存在: {'✅' if os.path.exists(gispro_python) else '❌'}"
    ])

    messages, missing = collect_config_status(config)
    print_info("📁 数据文件", messages)

    # 列出预测属性栅格
    rf_dir = config["rf_dir"]
    if os.path.isdir(rf_dir):
        tif_files = [f for f in os.listdir(rf_dir) if f.lower().endswith('.tif')]
        if tif_files:
            print_info("🗺️ 发现的属性栅格", tif_files)
        else:
            missing.append("预测属性栅格目录中没有 tif 文件")

    print_info("⚙️ 运行流程", [
        "1. 残差克里金插值 + 叠加校正",
        "2. 分级栅格生成 + 精度对比",
        "3. 验证集点表导出",
        "4. 训练集点表导出",
        "5. 补充统计表（表1、表2、表3）生成"
    ])

    print_separator()
    print("\n⚠️ 注意：")
    print("  • 确保所有路径配置正确")
    print("  • 确保 ArcGIS Pro 已正确安装并授权")
    print("  • 运行过程中请勿关闭 ArcGIS Pro（如果正在运行）")
    print_separator()

    if missing:
        print("\n❌ 当前不能开始执行，还缺这些输入：")
        for item in missing:
            print(f"  • {item}")
        print("\n请先补齐缺失输入，或修改脚本中的默认配置后再运行。")
        return False

    if check_only:
        print("\n✅ 输入资料存在性检查通过。当前为 --check-only 模式，不执行处理脚本。")
        return False

    while True:
        response = input("\n❓ 确认配置正确并开始运行？(y/n): ").strip().lower()
        if response in ['y', 'yes', '是']:
            return True
        elif response in ['n', 'no', '否']:
            return False
        else:
            print("  请输入 y/n 或 是/否")


def run_script(propy_path, script_name, description):
    """运行单个脚本"""
    print(f"\n{'=' * 80}")
    print(f"▶️  开始: {description}")
    print(f"{'=' * 80}")
    print(f"   脚本: {script_name}")
    print()

    cmd = [propy_path, script_name]
    try:
        result = subprocess.run(cmd, cwd=os.path.dirname(os.path.abspath(__file__)))
        if result.returncode == 0:
            print(f"\n✅ {description} - 完成")
            return True
        else:
            print(f"\n❌ {description} - 失败（返回码: {result.returncode}）")
            return False
    except Exception as e:
        print(f"\n❌ {description} - 异常: {str(e)}")
        return False


def main():
    """主函数"""
    check_only = "--check-only" in sys.argv

    if check_only:
        confirm_config(check_only=True)
        return

    # 显示配置并确认
    if not confirm_config():
        print("\n❌ 用户取消运行")
        return

    # 路径配置
    gispro_propy = build_default_config()["gispro_propy"]
    scripts_dir = os.path.dirname(os.path.abspath(__file__))

    scripts = [
        ("rf_residual_kriging_pipeline.py", "残差克里金叠加校正与精度/分级对比"),
        ("export_corrected_validation_points.py", "验证集点表导出"),
        ("export_training_points.py", "训练集点表导出"),
        ("landuse_township_tables.py", "补充统计表生成"),
    ]

    # 逐个运行脚本
    success_count = 0
    for script_name, description in scripts:
        script_path = os.path.join(scripts_dir, script_name)
        if not os.path.exists(script_path):
            print(f"\n⚠️  跳过: {description}（脚本不存在: {script_path}）")
            continue
        if run_script(gispro_propy, script_path, description):
            success_count += 1

    # 总结
    print(f"\n{'=' * 80}")
    print("📊 运行总结")
    print(f"{'=' * 80}")
    print(f"   成功: {success_count}/{len(scripts)} 个流程")
    if success_count == len(scripts):
        print("✅ 所有流程已完成！")
    else:
        print(f"⚠️  部分流程未完成，请检查")


if __name__ == "__main__":
    main()

"""
送货单对账单生成工具 - 简化版（无GUI依赖）
"""
import os
from pathlib import Path
from merge_delivery_orders import merge_delivery_orders, create_statement
import pandas as pd


def main():
    print("=" * 60)
    print("           送货单对账单生成工具")
    print("=" * 60)
    print()

    # 获取原始数据路径
    default_raw = "raw-data"
    raw_data_dir = input(f"请输入原始数据文件夹路径 (默认: {default_raw}): ").strip()
    if not raw_data_dir:
        raw_data_dir = default_raw

    # 检查路径是否存在
    if not os.path.exists(raw_data_dir):
        print(f"❌ 错误: 文件夹不存在: {raw_data_dir}")
        input("按回车键退出...")
        return

    # 获取输出路径
    default_output = "output"
    output_dir = input(f"请输入输出文件夹路径 (默认: {default_output}): ").strip()
    if not output_dir:
        output_dir = default_output

    print()
    print("-" * 60)
    print(f"原始数据文件夹: {raw_data_dir}")
    print(f"输出文件夹: {output_dir}")
    print("-" * 60)
    print()

    confirm = input("是否开始处理? (y/n): ").strip().lower()
    if confirm != 'y':
        print("已取消")
        return

    print()
    print("🚀 开始处理...")
    print()

    try:
        # 创建输出目录
        Path(output_dir).mkdir(parents=True, exist_ok=True)

        # 合并送货单
        output_file = os.path.join(output_dir, 'merged_delivery_orders.xlsx')
        print("📊 正在合并送货单数据...")
        df_summary = merge_delivery_orders(
            raw_data_dir=raw_data_dir,
            output_file=output_file
        )

        # 读取详细数据用于生成对账单
        df_all = pd.read_excel(output_file, sheet_name='详细数据')

        # 转换日期列为datetime类型
        df_all['日期'] = pd.to_datetime(df_all['日期'])

        # 提取年月
        df_all['年月'] = df_all['日期'].dt.to_period('M')

        # 按客户和年月分组
        grouped = df_all.groupby(['客户', '年月'])

        print(f"\n📝 开始生成对账单...")
        print(f"共有 {len(grouped)} 个客户月份组合\n")

        # 创建输出目录
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)

        # 为每个客户的每个月生成对账单
        skipped_count = 0
        generated_count = 0

        for (customer, year_month), group_data in grouped:
            # 创建客户文件夹
            customer_dir = output_path / customer
            customer_dir.mkdir(exist_ok=True)

            # 生成文件名
            statement_file = customer_dir / f'statement_{customer}_{year_month}.xlsx'

            # 检查文件是否已存在
            if statement_file.exists():
                print(f"⏭️  对账单已存在，跳过: {statement_file.name}")
                skipped_count += 1
                continue

            # 格式化年月显示
            year_month_str = f'{year_month.year}年{year_month.month}月'

            # 生成对账单
            create_statement(
                group_data,
                customer_name=customer,
                year_month=year_month_str,
                output_file=str(statement_file)
            )
            generated_count += 1

        print()
        print("=" * 60)
        print("✅ 所有对账单生成完成！")
        print(f"新生成: {generated_count} 个对账单")
        print(f"已跳过: {skipped_count} 个对账单")
        print(f"文件保存位置: {output_dir}")
        print("=" * 60)

    except Exception as e:
        print()
        print("=" * 60)
        print(f"❌ 错误: {str(e)}")
        print("=" * 60)
        import traceback
        traceback.print_exc()

    print()
    input("按回车键退出...")


if __name__ == '__main__':
    main()

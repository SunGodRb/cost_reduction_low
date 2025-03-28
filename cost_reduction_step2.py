#%%

print("""
      
      
      
 ██████╗ ██████╗ ███████╗████████╗                ██████╗ ███████╗██████╗ ██╗   ██╗ ██████╗████████╗██╗ ██████╗ ███╗   ██╗
██╔════╝██╔═══██╗██╔════╝╚══██╔══╝                ██╔══██╗██╔════╝██╔══██╗██║   ██║██╔════╝╚══██╔══╝██║██╔═══██╗████╗  ██║
██║     ██║   ██║███████╗   ██║                   ██████╔╝█████╗  ██║  ██║██║   ██║██║        ██║   ██║██║   ██║██╔██╗ ██║
██║     ██║   ██║╚════██║   ██║                   ██╔══██╗██╔══╝  ██║  ██║██║   ██║██║        ██║   ██║██║   ██║██║╚██╗██║
╚██████╗╚██████╔╝███████║   ██║       ███████╗    ██║  ██║███████╗██████╔╝╚██████╔╝╚██████╗   ██║   ██║╚██████╔╝██║ ╚████║
 ╚═════╝ ╚═════╝ ╚══════╝   ╚═╝       ╚══════╝    ╚═╝  ╚═╝╚══════╝╚═════╝  ╚═════╝  ╚═════╝   ╚═╝   ╚═╝ ╚═════╝ ╚═╝  ╚═══╝
                                                                                                                          
                             ███████╗████████╗███████╗██████╗     ████████╗██╗    ██╗ ██████╗                                  
                             ██╔════╝╚══██╔══╝██╔════╝██╔══██╗    ╚══██╔══╝██║    ██║██╔═══██╗                                 
                             ███████╗   ██║   █████╗  ██████╔╝       ██║   ██║ █╗ ██║██║   ██║                                 
                             ╚════██║   ██║   ██╔══╝  ██╔═══╝        ██║   ██║███╗██║██║   ██║                                 
                             ███████║   ██║   ███████╗██║            ██║   ╚███╔███╔╝╚██████╔╝                                 
                             ╚══════╝   ╚═╝   ╚══════╝╚═╝            ╚═╝    ╚══╝╚══╝  ╚═════╝                                  
                                                                                                                          

      
      
      """)
input("请确认processing文件夹中的文件已经更新完毕，按回车键继续,否则直接退出")
weekn = input('当前时间为wk几？，输入格式为wk1,wk2,wk3,wk4  \n')
current_week = int(weekn.replace('wk', ''))
monthn = input('当前时间为几月？，输入格式为1,2,3,4,5,6,7,8,9,10,11,12  \n')
current_month = int(monthn)
# 分割线
def print_section(title, char='-', length=30):
    print(f"{title}")
    print(char * length)
from datetime import datetime
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')
#%%
raw_df = pd.read_excel(r'processing\机型-F物料.xlsx')
product_price_df = pd.read_excel(r'input\整机基准价和历史价.xlsx')
print('正在对raw_df进行汇总处理...')

# 定义需要保留第一个值的列
first_value_cols = ['物料描述', '系列', '项目号', 'PC', '产品', '版本', '渠道', '品牌', '工厂', '状态']

# 定义需要求和的周物料成本列
week_cost_cols = [f'wk{i}-物料成本' for i in range(1, 5)]
week_bom_cols = [f'bom成本-wk{i}' for i in range(1, 5)]

# 定义需要求最大值的月整机数量列
month_quantity_cols = [f'{i}月整机数量' for i in range(1, 13)]
month_forecast_cols = [f'{i}月整机预测量' for i in range(1, 13)]

# 定义需要求和的月目标bom成本列
month_bom_cols = [f'{i}月目标bom成本' for i in range(1, 13)]
month_cost_cols = [f'{i}月物料成本' for i in range(1, 13)]

# 创建聚合字典
agg_dict = {}

# 1. 对first_value_cols中的列保留第一个值
for col in first_value_cols:
    if col in raw_df.columns:
        agg_dict[col] = 'first'

# 2. 对各周的物料成本进行求和
for week_cost, week_bom in zip(week_cost_cols, week_bom_cols):
    if week_bom in raw_df.columns:
        agg_dict[week_bom] = 'sum'

# 3. 对各月整机数量取最大值
for month_quantity, month_forecast in zip(month_quantity_cols, month_forecast_cols):
    if month_forecast in raw_df.columns:
        agg_dict[month_forecast] = 'max'

# 4. 对各月目标bom成本进行求和
for month_bom, month_cost in zip(month_bom_cols, month_cost_cols):
    if month_bom in raw_df.columns:
        agg_dict[month_bom] = 'sum'

# 5. 对最终物料成本进行求和
if '最终成本(目标价&财务核价取低)' in raw_df.columns:
    agg_dict['最终成本(目标价&财务核价取低)'] = 'sum'

# 执行分组聚合
result_df = raw_df.groupby('整机编码').agg(agg_dict).reset_index()

# 重命名列
rename_dict = {}

# 重命名月整机预测量为月整机数量
for month_forecast, month_quantity in zip(month_forecast_cols, month_quantity_cols):
    if month_forecast in result_df.columns:
        rename_dict[month_forecast] = month_quantity

# 重命名bom成本为周物料成本
for week_bom, week_cost in zip(week_bom_cols, week_cost_cols):
    if week_bom in result_df.columns:
        rename_dict[week_bom] = week_cost

# 重命名月目标bom成本为月物料成本
for month_bom, month_cost in zip(month_bom_cols, month_cost_cols):
    if month_bom in result_df.columns:
        rename_dict[month_bom] = month_cost

# 重命名最终物料成本
if '最终成本(目标价&财务核价取低)' in result_df.columns:
    rename_dict['最终成本(目标价&财务核价取低)'] = '最终物料成本(目标&财务核价取低)'

result_df = result_df.rename(columns=rename_dict)

print('正在删除小于当前月的物料成本列...')

# 获取所有列名
all_columns = result_df.columns.tolist()

# 找出需要删除的列（小于当前月的物料成本列）
columns_to_drop = []
for col in all_columns:
    if '月物料成本' in col:
        try:
            # 从列名中提取月份数字
            month_num = int(''.join(filter(str.isdigit, col.split('月')[0])))
            if month_num < current_month:
                columns_to_drop.append(col)
        except:
            continue

# 删除小于当前月的物料成本列
result_df = result_df.drop(columns=columns_to_drop)
#%%
print('正在处理product_price_df...')

# 保留基础列
base_columns = ['整机编码', '基准生产成本', '基准制费成本', '基准物料成本']

# 获取所有列名
all_columns = product_price_df.columns.tolist()

# 找出需要删除的列（大于等于当前月的列）
columns_to_drop = []
for col in all_columns:
    if col not in base_columns:  # 如果不是基础列
        try:
            # 尝试从列名中提取月份数字
            month_num = int(''.join(filter(str.isdigit, col.split('月')[0])))
            if month_num >= current_month:
                columns_to_drop.append(col)
        except:
            continue

# 删除大于等于当前月的列
product_price_df = product_price_df.drop(columns=columns_to_drop)

print('正在连接result_df和product_price_df...')

# 左连接result_df和product_price_df
result_df = result_df.merge(product_price_df, on='整机编码', how='left')


#%%
print('正在补充未来月的制费成本...')

# 获取所有列名
all_columns = result_df.columns.tolist()

# 找出制费成本列
fee_columns = []
for col in all_columns:
    if '月制费成本' in col:
        try:
            month_num = int(''.join(filter(str.isdigit, col.split('月')[0])))
            fee_columns.append((month_num, col))
        except:
            continue

# 按月份排序
fee_columns.sort(key=lambda x: x[0])

# 补充未来月的制费成本
for i in range(current_month, 7):  # 从当前月补充到6月
    target_col = f'{i}月制费成本'
    if target_col not in result_df.columns:
        # 找到上一个月的制费成本列
        prev_month = i - 1
        prev_col = f'{prev_month}月制费成本'
        if prev_col in result_df.columns:
            # 处理空值：如果上一月的值为空，则设为0
            result_df[target_col] = result_df[prev_col].fillna(0)

print('制费成本补充完成！')
#%%
print('正在补充未来月的生产成本...')

# 补充未来月的生产成本
for i in range(current_month, 7):  # 从当前月补充到6月
    target_col = f'{i}月生产成本'
    fee_col = f'{i}月制费成本'
    material_col = f'{i}月物料成本'
    
    # 确保制费成本和物料成本列都存在
    if fee_col in result_df.columns and material_col in result_df.columns:
        # 处理空值：如果任一值为空，则设为0
        fee_values = result_df[fee_col].fillna(0)
        material_values = result_df[material_col].fillna(0)
        result_df[target_col] = fee_values + material_values

print('未来月生产成本补充完成！')

print('正在计算各月生产相关金额...')

# 计算各月生产相关金额
for i in range(1, 7):  # 计算1月到6月
    # 1. 计算各月生产金额
    cost_col = f'{i}月生产成本'
    quantity_col = f'{i}月整机数量'
    production_col = f'{i}月生产金额'
    
    if cost_col in result_df.columns and quantity_col in result_df.columns:
        # 处理空值：如果任一值为空，则设为0
        cost_values = result_df[cost_col].fillna(0)
        quantity_values = result_df[quantity_col].fillna(0)
        result_df[production_col] = cost_values * quantity_values
    
    # 2. 计算各月基准生产金额
    base_cost_col = '基准生产成本'
    base_production_col = f'{i}月基准生产金额'
    
    if base_cost_col in result_df.columns and quantity_col in result_df.columns:
        # 处理空值：如果任一值为空，则设为0
        base_cost_values = result_df[base_cost_col].fillna(0)
        quantity_values = result_df[quantity_col].fillna(0)
        result_df[base_production_col] = base_cost_values * quantity_values
    
    # 3. 计算各月生产降本金额
    saving_col = f'{i}月生产降本金额'
    
    if base_production_col in result_df.columns and production_col in result_df.columns:
        # 使用apply函数处理条件计算
        result_df[saving_col] = result_df.apply(
            lambda row: row[base_production_col] - row[production_col] 
            if row[base_production_col] != 0 and row[production_col] != 0 
            else 0, 
            axis=1
        )
    
    # 4. 计算各月降本率
    rate_col = f'{i}月生产降本率'
    
    if saving_col in result_df.columns and base_production_col in result_df.columns:
        # 使用apply函数处理条件计算
        result_df[rate_col] = result_df.apply(
            lambda row: row[saving_col] / row[base_production_col] 
            if row[base_production_col] != 0 
            else 0, 
            axis=1
        )

print('各月生产相关金额计算完成！')
#%%
print('正在计算各月物料相关金额...')

# 计算各月物料相关金额
for i in range(1, 7):  # 计算1月到6月
    # 1. 计算各月物料金额
    cost_col = f'{i}月物料成本'
    quantity_col = f'{i}月整机数量'
    production_col = f'{i}月物料金额'
    
    if cost_col in result_df.columns and quantity_col in result_df.columns:
        # 处理空值：如果任一值为空，则设为0
        cost_values = result_df[cost_col].fillna(0)
        quantity_values = result_df[quantity_col].fillna(0)
        result_df[production_col] = cost_values * quantity_values
    
    # 2. 计算各月基准物料金额
    base_cost_col = '基准物料成本'
    base_production_col = f'{i}月基准物料金额'
    
    if base_cost_col in result_df.columns and quantity_col in result_df.columns:
        # 处理空值：如果任一值为空，则设为0
        base_cost_values = result_df[base_cost_col].fillna(0)
        quantity_values = result_df[quantity_col].fillna(0)
        result_df[base_production_col] = base_cost_values * quantity_values
    
    # 3. 计算各月物料降本金额
    saving_col = f'{i}月物料降本金额'
    
    if base_production_col in result_df.columns and production_col in result_df.columns:
        # 使用apply函数处理条件计算
        result_df[saving_col] = result_df.apply(
            lambda row: row[base_production_col] - row[production_col] 
            if row[base_production_col] != 0 and row[production_col] != 0 
            else 0, 
            axis=1
        )
    
    # 4. 计算各月物料降本率
    rate_col = f'{i}月物料降本率'
    
    if saving_col in result_df.columns and base_production_col in result_df.columns:
        # 使用apply函数处理条件计算
        result_df[rate_col] = result_df.apply(
            lambda row: row[saving_col] / row[base_production_col] 
            if row[base_production_col] != 0 
            else 0, 
            axis=1
        )

print('各月物料相关金额计算完成！')
#%%
print('正在计算各月制费相关金额...')

# 计算各月制费相关金额
for i in range(1, 7):  # 计算1月到6月
    # 1. 计算各月制费金额
    cost_col = f'{i}月制费成本'
    quantity_col = f'{i}月整机数量'
    production_col = f'{i}月制费金额'
    
    if cost_col in result_df.columns and quantity_col in result_df.columns:
        # 处理空值：如果任一值为空，则设为0
        cost_values = result_df[cost_col].fillna(0)
        quantity_values = result_df[quantity_col].fillna(0)
        result_df[production_col] = cost_values * quantity_values
    
    # 2. 计算各月基准制费金额
    base_cost_col = '基准制费成本'
    base_production_col = f'{i}月基准制费金额'
    
    if base_cost_col in result_df.columns and quantity_col in result_df.columns:
        # 处理空值：如果任一值为空，则设为0
        base_cost_values = result_df[base_cost_col].fillna(0)
        quantity_values = result_df[quantity_col].fillna(0)
        result_df[base_production_col] = base_cost_values * quantity_values
    
    # 3. 计算各月制费降本金额
    saving_col = f'{i}月制费降本金额'
    
    if base_production_col in result_df.columns and production_col in result_df.columns:
        # 使用apply函数处理条件计算
        result_df[saving_col] = result_df.apply(
            lambda row: row[base_production_col] - row[production_col] 
            if row[base_production_col] != 0 and row[production_col] != 0 
            else 0, 
            axis=1
        )
    
    # 4. 计算各月制费降本率
    rate_col = f'{i}月制费降本率'
    
    if saving_col in result_df.columns and base_production_col in result_df.columns:
        # 使用apply函数处理条件计算
        result_df[rate_col] = result_df.apply(
            lambda row: row[saving_col] / row[base_production_col] 
            if row[base_production_col] != 0 
            else 0, 
            axis=1
        )

print('各月制费相关金额计算完成！')

print('正在计算最终物料成本-降本额...')

# 计算最终物料成本-降本额
final_cost_col = '最终物料成本(目标&财务核价取低)'
quantity_col = '6月整机数量'
target_col = '最终物料成本-降本额'

if final_cost_col in result_df.columns and quantity_col in result_df.columns:
    # 处理空值：如果任一值为空，则设为0
    final_cost_values = result_df[final_cost_col].fillna(0)
    quantity_values = result_df[quantity_col].fillna(0)
    result_df[target_col] = final_cost_values * quantity_values

print('最终物料成本-降本额计算完成！')
#%%
print('正在计算各周物料降本额...')

# 获取当前月整机数量列名
current_month_quantity_col = f'{current_month}月整机数量'

# 计算各周物料降本额
for week in range(1, 5):
    # 定义相关列名
    week_cost_col = f'wk{week}-物料成本'
    target_col = f'wk{week}-物料降本额'
    
    # 检查必要的列是否存在
    if week_cost_col in result_df.columns and '基准物料成本' in result_df.columns and current_month_quantity_col in result_df.columns:
        # 处理空值：如果任一值为空，则设为0
        base_cost_values = result_df['基准物料成本'].fillna(0)
        week_cost_values = result_df[week_cost_col].fillna(0)
        quantity_values = result_df[current_month_quantity_col].fillna(0)
        
        # 计算物料降本额
        result_df[target_col] = result_df.apply(
            lambda row: (row['基准物料成本'] - row[week_cost_col]) * row[current_month_quantity_col]
            if row['基准物料成本'] != 0 and row[week_cost_col] != 0 and row[current_month_quantity_col] != 0
            else 0,
            axis=1
        )

print('各周物料降本额计算完成！')
#%%
print('正在重命名列...')

# 获取所有列名
all_columns = result_df.columns.tolist()

# 创建重命名字典
rename_dict = {}

# 遍历所有列名
for col in all_columns:
    # 检查列名是否包含月份
    if '月' in col:
        try:
            # 从列名中提取月份数字
            month_num = int(''.join(filter(str.isdigit, col.split('月')[0])))
            # 根据月份和current_month的关系决定后缀
            suffix = '-实际' if month_num > current_month else '-预估'
            # 如果列名还没有后缀，则添加后缀
            if not col.endswith('-实际') and not col.endswith('-预估'):
                rename_dict[col] = col + suffix
        except:
            continue

# 执行重命名
result_df = result_df.rename(columns=rename_dict)

print('列重命名完成！')

print('正在调整列顺序...')

# 定义新的列顺序
new_column_order = [
    '整机编码', '物料描述', '系列', '项目号', 'PC', '产品', '版本', '渠道', '品牌', '工厂', '状态',
    '基准生产成本', '基准制费成本', '基准物料成本', '最终物料成本(目标&财务核价取低)',
    'wk1-物料成本', 'wk2-物料成本', 'wk3-物料成本', 'wk4-物料成本',
    '最终物料成本-降本额', 'wk1-物料降本额', 'wk2-物料降本额', 'wk3-物料降本额', 'wk4-物料降本额'
]

# 添加1-6月的列
for month in range(1, 7):
    suffix = '-实际' if month <= current_month else '-预估'
    month_columns = [
        f'{month}月整机数量{suffix}',
        f'{month}月生产成本{suffix}',
        f'{month}月生产金额{suffix}',
        f'{month}月基准生产金额{suffix}',
        f'{month}月生产降本金额{suffix}',
        f'{month}月生产降本率{suffix}',
        f'{month}月制费成本{suffix}',
        f'{month}月制费金额{suffix}',
        f'{month}月基准制费金额{suffix}',
        f'{month}月制费降本金额{suffix}',
        f'{month}月制费降本率{suffix}',
        f'{month}月物料成本{suffix}',
        f'{month}月物料金额{suffix}',
        f'{month}月基准物料金额{suffix}',
        f'{month}月物料降本金额{suffix}',
        f'{month}月物料降本率{suffix}'
    ]
    new_column_order.extend(month_columns)

# 检查所有列是否存在
existing_columns = [col for col in new_column_order if col in result_df.columns]
missing_columns = [col for col in new_column_order if col not in result_df.columns]

if missing_columns:
    print(f"警告：以下列不存在：{missing_columns}")

# 重新排列列
result_df = result_df[existing_columns]

print('列顺序调整完成！')
#%%
current_time = datetime.now()
time_suffix = current_time.strftime("%Y%m%d_%H%M")
print('正在生成整机降本文件...')
result_df.to_excel(r'output\整机降本{}.xlsx'.format(time_suffix),index=False)
print('整机降本文件生成完成！')
input("""

            ______ _   __ ____              
           / ____// | / // __ \             
 ______   / __/  /  |/ // / / /  ______     
/_____/  / /___ / /|  // /_/ /  /_____/     
        /_____//_/ |_//_____/               
                                            

"""
)
# %%

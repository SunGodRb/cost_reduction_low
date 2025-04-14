#%%
print(''' 



 ██████╗ ██████╗ ███████╗████████╗                ██████╗ ███████╗██████╗ ██╗   ██╗ ██████╗████████╗██╗ ██████╗ ███╗   ██╗
██╔════╝██╔═══██╗██╔════╝╚══██╔══╝                ██╔══██╗██╔════╝██╔══██╗██║   ██║██╔════╝╚══██╔══╝██║██╔═══██╗████╗  ██║
██║     ██║   ██║███████╗   ██║                   ██████╔╝█████╗  ██║  ██║██║   ██║██║        ██║   ██║██║   ██║██╔██╗ ██║
██║     ██║   ██║╚════██║   ██║                   ██╔══██╗██╔══╝  ██║  ██║██║   ██║██║        ██║   ██║██║   ██║██║╚██╗██║
╚██████╗╚██████╔╝███████║   ██║       ███████╗    ██║  ██║███████╗██████╔╝╚██████╔╝╚██████╗   ██║   ██║╚██████╔╝██║ ╚████║
 ╚═════╝ ╚═════╝ ╚══════╝   ╚═╝       ╚══════╝    ╚═╝  ╚═╝╚══════╝╚═════╝  ╚═════╝  ╚═════╝   ╚═╝   ╚═╝ ╚═════╝ ╚═╝  ╚═══╝
                                                                                                                          
                            ███████╗████████╗███████╗██████╗        ██████╗ ███╗   ██╗███████╗                                        
                            ██╔════╝╚══██╔══╝██╔════╝██╔══██╗      ██╔═══██╗████╗  ██║██╔════╝                                        
                            ███████╗   ██║   █████╗  ██████╔╝█████╗██║   ██║██╔██╗ ██║█████╗                                          
                            ╚════██║   ██║   ██╔══╝  ██╔═══╝ ╚════╝██║   ██║██║╚██╗██║██╔══╝                                          
                            ███████║   ██║   ███████╗██║           ╚██████╔╝██║ ╚████║███████╗                                        
                            ╚══════╝   ╚═╝   ╚══════╝╚═╝            ╚═════╝ ╚═╝  ╚═══╝╚══════╝                                        
                                                                                                                          



''')
# 分割线
def print_section(title, char='-', length=30):
    print(f"{title}")
    print(char * length)
from datetime import datetime
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

weekn = input('当前时间为wk几？，输入格式为wk1,wk2,wk3,wk4  \n')
current_week = int(weekn.replace('wk', ''))
monthn = input('当前时间为几月？，输入格式为1,2,3,4,5,6,7,8,9,10,11,12  \n')
current_month = int(monthn)
print('正在读取整机数量清单...')
product_forecast = pd.read_excel(r'input\整机数量清单.xlsx').rename(columns={'物料号':'整机编码'}).drop(columns=['颜','线体','MP首单时间'])
#%%
print('正在读取BOM...')
use_cols = ['物料编码','中文名称','0','单位','基本用量计算组件数量','采购类型','本币单价','BOM成本','供应商描述','是否暂估价','散装物料(BOM)','散装物料']
bom = pd.read_excel(r'input\BOM.xlsx',usecols=use_cols)
bom = bom.rename(columns={'基本用量计算组件数量':'BOM用量','供应商描述':'供应商',
                          '是否暂估价':'价格类型','本币单价':'上月单价','BOM成本':'上月成本'})
bom_info = bom.copy()
# 在最左侧添加整机编码列
print('正在添加整机编码列...')
bom['BOM用量'] = pd.to_numeric(bom['BOM用量'], errors='coerce').fillna(0)
bom['上月成本'] = pd.to_numeric(bom['上月成本'], errors='coerce').fillna(0)
bom['上月单价'] = pd.to_numeric(bom['上月单价'], errors='coerce').fillna(0)
bom.insert(0, '整机编码', np.nan)
mask = (bom['0'] == 0) | (bom['0'] == '0')
bom.loc[mask, '整机编码'] = bom.loc[mask, '物料编码']
bom['整机编码'] = bom['整机编码'].fillna(method='ffill')
bom = bom.drop(columns=['0'])
#%%
# 找出product_forecast中存在但不在bom中的整机编码
print("\n正在查找缺失BOM的整机编码...")
forecast_machine_codes = set(product_forecast['整机编码'].unique())
bom_machine_codes = set(bom['整机编码'].unique())
missing_machine_codes = forecast_machine_codes - bom_machine_codes
if missing_machine_codes:
    missing_machines_info = product_forecast[product_forecast['整机编码'].isin(missing_machine_codes)][['整机编码', '物料描述']].drop_duplicates()
    print(f"\n以下{len(missing_machine_codes)}个整机编码在需求预测中存在，但在BOM中缺失：")
    for _, row in missing_machines_info.iterrows():
        print(f"- {row['整机编码']} ({row['物料描述']})")
        input("\n是否继续执行程序，按回车继续，点击×退出")
else:
    print("\n所有需求预测中的整机编码都在BOM中存在。")
#%%
print('正在读取最低基准价...')
min_price = pd.read_excel(r'input\大表基准价.xlsx')
min_price = min_price[['SAP物料编码','最低价基价']].drop_duplicates(subset=['SAP物料编码'])
min_price.rename(columns={'SAP物料编码':'物料编码','最低价基价':'最低基价(含还原)'},inplace=True)

print('正在读取物料主数据...')
master_data = pd.read_excel(r'input\主数据.xlsx')
master_data = master_data[['物料号','三级分类','二级分类','一级分类','采购经理','采购']].rename(columns={'物料号':'物料编码','三级分类':'小分类','二级分类':'中分类'}).drop_duplicates(subset=['物料编码'])

print('正在读取财务核价')
financial_price = pd.read_excel(r'input\财务核价物料清单.xlsx')
financial_price = financial_price[['SAP物料编码','财务核价']].rename(columns={'SAP物料编码':'物料编码'}).drop_duplicates(subset=['物料编码'])

#%%
print('正在读取大表最低价')
big_table_price1 = pd.read_excel(r'input\大表价格-wk1.xlsx')
big_table_price1 = big_table_price1.rename(columns={'SAP物料编码':'物料编码'})
#%%
# 处理big_table_price1的筛选逻辑
print('正在处理大表价格-wk1的筛选...')
# 重置价格
big_table_price1 = big_table_price1.drop(columns=['价格']).rename(columns={'最终价格(含税人民币)':'价格'})
# 只保留供应商编码长度不为4位的，或者供应商编码=1200的
big_table_price1 = big_table_price1[(big_table_price1['供应商编码'].astype(str).str.len() != 4) | (big_table_price1['供应商编码'].astype(str) == '1200')]
# 1. 筛选价格在有效期内的记录
#%%
today = pd.Timestamp.now().date()
# 处理日期转换，保留9999-12-31的日期
big_table_price1['有效期开始时间'] = pd.to_datetime(big_table_price1['有效期开始时间'], errors='coerce').dt.date
# 对于9999-12-31的日期，先转换为字符串再处理
big_table_price1['有效期到期时间'] = big_table_price1['有效期到期时间'].astype(str)
big_table_price1.loc[big_table_price1['有效期到期时间'] == '9999-12-31', '有效期到期时间'] = pd.Timestamp('9999-12-31').date()
big_table_price1.loc[big_table_price1['有效期到期时间'] != '9999-12-31', '有效期到期时间'] = pd.to_datetime(big_table_price1['有效期到期时间'], errors='coerce').dt.date

# 特殊处理9999-12-31的日期
valid_price = big_table_price1[
    ((big_table_price1['有效期开始时间'] <= today) & 
    ((big_table_price1['有效期到期时间'] >= today) | 
     (big_table_price1['有效期到期时间'] == pd.Timestamp('9999-12-31').date())))
].copy()

# 2. 按物料编码分组，优先取正式价格，其次取暂估价，都取创建日期最新的
valid_price['创建日期'] = pd.to_datetime(valid_price['创建日期'], errors='coerce')
valid_price = valid_price.sort_values('创建日期', ascending=False).drop_duplicates(subset=['物料编码','价格类型'], keep='first')  # 按创建日期降序排序

# 先尝试获取正式价格
formal_price = valid_price[valid_price['价格类型'] == '正式价'].groupby('物料编码').first().reset_index()
# 对于没有正式价格的物料，获取暂估价
temp_price = valid_price[valid_price['价格类型'] == '试产价'].groupby('物料编码').first().reset_index()
# 合并两种价格
final_price = pd.concat([formal_price, temp_price])
# 重命名价格列
final_price = final_price.rename(columns={'价格': '最低有效价-wk1'})
# 只保留需要的列
big_table_price1 = final_price[['物料编码','价格类型','最低有效价-wk1']].sort_values(by=['物料编码','价格类型','最低有效价-wk1'], ascending=True).drop_duplicates(subset=['物料编码'],keep='first').drop(columns=['价格类型'])

print('正在读取大表最低价')
big_table_price2 = pd.read_excel(r'input\大表价格-wk2.xlsx')
big_table_price2 = big_table_price2.rename(columns={'SAP物料编码':'物料编码'})

# 处理big_table_price2的筛选逻辑
print('正在处理大表价格-wk2的筛选...')
# 重置价格
big_table_price2 = big_table_price2.drop(columns=['价格']).rename(columns={'最终价格(含税人民币)':'价格'})
# 只保留供应商编码长度不为4位的，或者供应商编码=1200的
big_table_price2 = big_table_price2[(big_table_price2['供应商编码'].astype(str).str.len() != 4) | (big_table_price2['供应商编码'].astype(str) == '1200')]
# 1. 筛选价格在有效期内的记录
big_table_price2['有效期开始时间'] = pd.to_datetime(big_table_price2['有效期开始时间'], errors='coerce').dt.date
# 对于9999-12-31的日期，先转换为字符串再处理
big_table_price2['有效期到期时间'] = big_table_price2['有效期到期时间'].astype(str)
big_table_price2.loc[big_table_price2['有效期到期时间'] == '9999-12-31', '有效期到期时间'] = pd.Timestamp('9999-12-31').date()
big_table_price2.loc[big_table_price2['有效期到期时间'] != '9999-12-31', '有效期到期时间'] = pd.to_datetime(big_table_price2['有效期到期时间'], errors='coerce').dt.date

# 特殊处理9999-12-31的日期
valid_price = big_table_price2[
    ((big_table_price2['有效期开始时间'] <= today) & 
    ((big_table_price2['有效期到期时间'] >= today) | 
     (big_table_price2['有效期到期时间'] == pd.Timestamp('9999-12-31').date())))
].copy()

# 2. 按物料编码分组，优先取正式价格，其次取暂估价，都取创建日期最新的
valid_price['创建日期'] = pd.to_datetime(valid_price['创建日期'], errors='coerce')
valid_price = valid_price.sort_values('创建日期', ascending=False).drop_duplicates(subset=['物料编码','价格类型'], keep='first')   # 按创建日期降序排序

# 先尝试获取正式价格
formal_price = valid_price[valid_price['价格类型'] == '正式价'].groupby('物料编码').first().reset_index()
# 对于没有正式价格的物料，获取暂估价
temp_price = valid_price[valid_price['价格类型'] == '试产价'].groupby('物料编码').first().reset_index()
# 合并两种价格
final_price = pd.concat([formal_price, temp_price])

# 重命名价格列
final_price = final_price.rename(columns={'价格': '最低有效价-wk2'})
# 只保留需要的列
big_table_price2 = final_price[['物料编码','价格类型','最低有效价-wk2']].sort_values(by=['物料编码','价格类型','最低有效价-wk2'], ascending=True).drop_duplicates(subset=['物料编码'],keep='first').drop(columns=['价格类型'])

print('正在读取大表最低价')
big_table_price3 = pd.read_excel(r'input\大表价格-wk3.xlsx')
big_table_price3 = big_table_price3.rename(columns={'SAP物料编码':'物料编码'})

# 处理big_table_price3的筛选逻辑
print('正在处理大表价格-wk3的筛选...')
# 重置价格
big_table_price3 = big_table_price3.drop(columns=['价格']).rename(columns={'最终价格(含税人民币)':'价格'})
# 只保留供应商编码长度不为4位的，或者供应商编码=1200的
big_table_price3 = big_table_price3[(big_table_price3['供应商编码'].astype(str).str.len() != 4) | (big_table_price3['供应商编码'].astype(str) == '1200')]
# 1. 筛选价格在有效期内的记录
big_table_price3['有效期开始时间'] = pd.to_datetime(big_table_price3['有效期开始时间'], errors='coerce').dt.date
# 对于9999-12-31的日期，先转换为字符串再处理
big_table_price3['有效期到期时间'] = big_table_price3['有效期到期时间'].astype(str)
big_table_price3.loc[big_table_price3['有效期到期时间'] == '9999-12-31', '有效期到期时间'] = pd.Timestamp('9999-12-31').date()
big_table_price3.loc[big_table_price3['有效期到期时间'] != '9999-12-31', '有效期到期时间'] = pd.to_datetime(big_table_price3['有效期到期时间'], errors='coerce').dt.date

# 特殊处理9999-12-31的日期
valid_price = big_table_price3[
    ((big_table_price3['有效期开始时间'] <= today) & 
    ((big_table_price3['有效期到期时间'] >= today) | 
     (big_table_price3['有效期到期时间'] == pd.Timestamp('9999-12-31').date())))
].copy()

# 2. 按物料编码分组，优先取正式价格，其次取暂估价，都取创建日期最新的
valid_price['创建日期'] = pd.to_datetime(valid_price['创建日期'], errors='coerce')
valid_price = valid_price.sort_values('创建日期', ascending=False).drop_duplicates(subset=['物料编码','价格类型'], keep='first')   # 按创建日期降序排序

# 先尝试获取正式价格
formal_price = valid_price[valid_price['价格类型'] == '正式价'].groupby('物料编码').first().reset_index()
# 对于没有正式价格的物料，获取暂估价
temp_price = valid_price[valid_price['价格类型'] == '试产价'].groupby('物料编码').first().reset_index()
# 合并两种价格
final_price = pd.concat([formal_price, temp_price])

# 重命名价格列
final_price = final_price.rename(columns={'价格': '最低有效价-wk3'})
# 只保留需要的列
big_table_price3 = final_price[['物料编码','价格类型','最低有效价-wk3']].sort_values(by=['物料编码','价格类型','最低有效价-wk3'], ascending=True).drop_duplicates(subset=['物料编码'],keep='first').drop(columns=['价格类型'])

print('正在读取大表最低价')
big_table_price4 = pd.read_excel(r'input\大表价格-wk4.xlsx')
big_table_price4 = big_table_price4.rename(columns={'SAP物料编码':'物料编码'})

# 处理big_table_price4的筛选逻辑
print('正在处理大表价格-wk4的筛选...')
# 重置价格
big_table_price4 = big_table_price4.drop(columns=['价格']).rename(columns={'最终价格(含税人民币)':'价格'})
big_table_price4 = big_table_price4[(big_table_price4['供应商编码'].astype(str).str.len() != 4) | (big_table_price4['供应商编码'].astype(str) == '1200')]
# 1. 筛选价格在有效期内的记录
big_table_price4['有效期开始时间'] = pd.to_datetime(big_table_price4['有效期开始时间'], errors='coerce').dt.date
# 对于9999-12-31的日期，先转换为字符串再处理
big_table_price4['有效期到期时间'] = big_table_price4['有效期到期时间'].astype(str)
big_table_price4.loc[big_table_price4['有效期到期时间'] == '9999-12-31', '有效期到期时间'] = pd.Timestamp('9999-12-31').date()
big_table_price4.loc[big_table_price4['有效期到期时间'] != '9999-12-31', '有效期到期时间'] = pd.to_datetime(big_table_price4['有效期到期时间'], errors='coerce').dt.date

# 特殊处理9999-12-31的日期
valid_price = big_table_price4[
    ((big_table_price4['有效期开始时间'] <= today) & 
    ((big_table_price4['有效期到期时间'] >= today) | 
     (big_table_price4['有效期到期时间'] == pd.Timestamp('9999-12-31').date())))
].copy()

# 2. 按物料编码分组，优先取正式价格，其次取暂估价，都取创建日期最新的
valid_price['创建日期'] = pd.to_datetime(valid_price['创建日期'], errors='coerce')
valid_price = valid_price.sort_values('创建日期', ascending=False).drop_duplicates(subset=['物料编码','价格类型'],keep='first')  # 按创建日期降序排序

# 先尝试获取正式价格
formal_price = valid_price[valid_price['价格类型'] == '正式价'].groupby('物料编码').first().reset_index()
# 对于没有正式价格的物料，获取暂估价
temp_price = valid_price[valid_price['价格类型'] == '试产价'].groupby('物料编码').first().reset_index()
# 合并两种价格
final_price = pd.concat([formal_price, temp_price])

# 重命名价格列
final_price = final_price.rename(columns={'价格': '最低有效价-wk4'})
# 只保留需要的列
big_table_price4 = final_price[['物料编码','价格类型','最低有效价-wk4']].sort_values(by=['物料编码','价格类型','最低有效价-wk4'], ascending=True).drop_duplicates(subset=['物料编码'],keep='first').drop(columns=['价格类型'])

print('所有大表价格处理完成！')
#%%
print('正在读取采购目标价')
purchase_target_price = pd.read_excel(r'input\采购目标价.xlsx')
purchase_target_price = purchase_target_price[['物料编码','1月目标价','2月目标价','3月目标价','4月目标价','5月目标价','6月目标价']].drop_duplicates(subset=['物料编码'])




#%%
product_bom = pd.merge(product_forecast, bom, on='整机编码', how='left')
product_bom = product_bom[(product_bom['采购类型']=='F')&(product_bom['散装物料(BOM)'].isna())&(product_bom['散装物料'].isna())]


# 定义需要汇总的列
groupby_cols = ['整机编码', '物料描述', '系列', '项目号', 'PC', '产品', '版本', '渠道', '品牌', '工厂', '状态', '物料编码']

# 找出需要求和的月份列（格式为"x月整机预测量"）
month_cols = [f"{i}月整机预测量" for i in range(1, 13) if f"{i}月整机预测量" in product_bom.columns]
# 将各月整机预测量转换为数值类型并填充0
for col in month_cols:
    product_bom[col] = pd.to_numeric(product_bom[col], errors='coerce').fillna(0)

# 定义聚合方法
agg_dict = {}
# 月份列求和
for col in month_cols:
    agg_dict[col] = 'sum'
# 上月成本和BOM用量求和
agg_dict.update({
    '上月成本': 'sum',
    'BOM用量': 'sum'
})
# 其他列保留第一个值
other_cols = [col for col in product_bom.columns if col not in groupby_cols + month_cols + ['上月成本', 'BOM用量']]
for col in other_cols:
    agg_dict[col] = 'first'

# 执行汇总操作
product_bom_summary = product_bom.groupby(groupby_cols).agg(agg_dict).reset_index()
#%%
# 计算各月物料预测量
for i in range(1, 13):
    month_col = f"{i}月整机预测量"
    if month_col in product_bom_summary.columns:
        material_forecast_col = f"{i}月物料预测量"
        product_bom_summary[material_forecast_col] = product_bom_summary[month_col] * product_bom_summary['BOM用量']
# %%
print('连接物料信息')
product_bom_summary = pd.merge(product_bom_summary, master_data, on='物料编码', how='left')
print('连接基准价')
product_bom_summary = pd.merge(product_bom_summary, min_price, on='物料编码', how='left')
product_bom_summary['基准bom成本'] = product_bom_summary['最低基价(含还原)']*product_bom_summary['BOM用量']
print('连接财务核价')
product_bom_summary = pd.merge(product_bom_summary, financial_price, on='物料编码', how='left')

print('连接各周大表最低价,并处理未来wk价格')
product_bom_summary = pd.merge(product_bom_summary, big_table_price1, on='物料编码', how='left')
product_bom_summary = pd.merge(product_bom_summary, big_table_price2, on='物料编码', how='left')
product_bom_summary = pd.merge(product_bom_summary, big_table_price3, on='物料编码', how='left')
product_bom_summary = pd.merge(product_bom_summary, big_table_price4, on='物料编码', how='left')
# 将各周的最低价转换为数值类型并填充0
print('将各周的最低价转换为数值类型并填充0...')
product_bom_summary['最低有效价-wk1'] = pd.to_numeric(product_bom_summary['最低有效价-wk1'], errors='coerce').fillna(0)
product_bom_summary['最低有效价-wk2'] = pd.to_numeric(product_bom_summary['最低有效价-wk2'], errors='coerce').fillna(0)
product_bom_summary['最低有效价-wk3'] = pd.to_numeric(product_bom_summary['最低有效价-wk3'], errors='coerce').fillna(0)
product_bom_summary['最低有效价-wk4'] = pd.to_numeric(product_bom_summary['最低有效价-wk4'], errors='coerce').fillna(0)

# 处理未来周的价格
print('处理未来周的价格...')

if current_week < 4:
    for week in range(current_week + 1, 5):
        price_col = f'最低有效价-wk{week}'
        if price_col in product_bom_summary.columns:
            product_bom_summary[price_col] = 0


print('连接采购目标价')
product_bom_summary = pd.merge(product_bom_summary, purchase_target_price, on='物料编码', how='left')
product_bom_summary['最终价(目标价&财务核价取低)'] = product_bom_summary[['财务核价','6月目标价']].fillna(float('inf')).min(axis=1).replace(float('inf'), pd.NA).fillna(0)
# 将各月目标价转换为数值类型并填充缺失值为0
target_price_cols = ['1月目标价','2月目标价','3月目标价','4月目标价','5月目标价','6月目标价']
for col in target_price_cols:
    product_bom_summary[col] = pd.to_numeric(product_bom_summary[col], errors='coerce').fillna(0)
# 计算wk1的BOM成本
print('计算wk1的BOM成本...')
product_bom_summary['bom成本-wk1'] = product_bom_summary.apply(
    lambda x: x['最低有效价-wk1'] * x['BOM用量'] if x['最低有效价-wk1'] > 0 else x['上月成本'],
    axis=1
)

# 计算wk2的BOM成本
print('计算wk2的BOM成本...')
product_bom_summary['bom成本-wk2'] = product_bom_summary.apply(
    lambda x: x['最低有效价-wk2'] * x['BOM用量'] if x['最低有效价-wk2'] > 0 else x['bom成本-wk1'],
    axis=1
)

# 计算wk3的BOM成本
print('计算wk3的BOM成本...')
product_bom_summary['bom成本-wk3'] = product_bom_summary.apply(
    lambda x: x['最低有效价-wk3'] * x['BOM用量'] if x['最低有效价-wk3'] > 0 else x['bom成本-wk2'],
    axis=1
)

# 计算wk4的BOM成本
print('计算wk4的BOM成本...')
product_bom_summary['bom成本-wk4'] = product_bom_summary.apply(
    lambda x: x['最低有效价-wk4'] * x['BOM用量'] if x['最低有效价-wk4'] > 0 else x['bom成本-wk3'],
    axis=1
)

# 计算1月目标BOM成本
print('计算1月目标BOM成本...')
product_bom_summary['1月目标bom成本'] = product_bom_summary.apply(
    lambda x: min(x['1月目标价'] * x['BOM用量'], x['bom成本-wk4']) if x['1月目标价'] > 0 else x['bom成本-wk4'],
    axis=1
)

# 计算2月目标BOM成本
print('计算2月目标BOM成本...')
product_bom_summary['2月目标bom成本'] = product_bom_summary.apply(
    lambda x: min(x['2月目标价'] * x['BOM用量'], x['1月目标bom成本']) if x['2月目标价'] > 0 else x['1月目标bom成本'],
    axis=1
)

# 计算3月目标BOM成本
print('计算3月目标BOM成本...')
product_bom_summary['3月目标bom成本'] = product_bom_summary.apply(
    lambda x: min(x['3月目标价'] * x['BOM用量'], x['2月目标bom成本']) if x['3月目标价'] > 0 else x['2月目标bom成本'],
    axis=1
)

# 计算4月目标BOM成本
print('计算4月目标BOM成本...')
product_bom_summary['4月目标bom成本'] = product_bom_summary.apply(
    lambda x: min(x['4月目标价'] * x['BOM用量'], x['3月目标bom成本']) if x['4月目标价'] > 0 else x['3月目标bom成本'],
    axis=1
)

# 计算5月目标BOM成本
print('计算5月目标BOM成本...')
product_bom_summary['5月目标bom成本'] = product_bom_summary.apply(
    lambda x: min(x['5月目标价'] * x['BOM用量'], x['4月目标bom成本']) if x['5月目标价'] > 0 else x['4月目标bom成本'],
    axis=1
)

# 计算6月目标BOM成本
print('计算6月目标BOM成本...')
product_bom_summary['6月目标bom成本'] = product_bom_summary.apply(
    lambda x: min(x['6月目标价'] * x['BOM用量'], x['5月目标bom成本']) if x['6月目标价'] > 0 else x['5月目标bom成本'],
    axis=1
)

# 计算最终成本(目标价&财务核价取低)
print('计算最终成本(目标价&财务核价取低)...')
product_bom_summary['最终成本(目标价&财务核价取低)'] = product_bom_summary.apply(
    lambda x: x['最终价(目标价&财务核价取低)'] * x['BOM用量'] if x['最终价(目标价&财务核价取低)'] > 0 else x['6月目标bom成本'],
    axis=1
)
#%%
# 新增列【系列+小分类】，列值为系列和小分类拼接
print('正在创建【系列+小分类】列...')
product_bom_summary['系列+小分类'] = product_bom_summary['系列'] + product_bom_summary['小分类']

# 新增列【项目+小分类】，列值为项目号、小分类拼接
print('正在创建【项目+小分类】列...')
product_bom_summary['项目+小分类'] = product_bom_summary['项目号'].str[:5] + product_bom_summary['小分类']

# 创建以系列+小分类为分组依据的最低价格表
print('正在创建系列+小分类最低价格表...')
series_min_price = product_bom_summary.groupby('系列+小分类').agg({
    f'最低有效价-wk{current_week}': 'min'
}).reset_index()

# 创建以项目+小分类为分组依据的最低价格表
print('正在创建项目+小分类最低价格表...')
project_min_price = product_bom_summary.groupby('项目+小分类').agg({
    f'最低有效价-wk{current_week}': 'min'
}).reset_index()

# 重命名列
series_min_price = series_min_price.rename(columns={f'最低有效价-wk{current_week}': f'同系列品类最低价'})
project_min_price = project_min_price.rename(columns={f'最低有效价-wk{current_week}': f'同项目品类最低价'})

# 连接两个最低价格表到product_bom_summary
print('正在连接最低价格表...')
product_bom_summary = pd.merge(product_bom_summary, series_min_price, on='系列+小分类', how='left')
product_bom_summary = pd.merge(product_bom_summary, project_min_price, on='项目+小分类', how='left')

# 计算同系列品类价差比
print('正在计算同系列品类价差比...')
product_bom_summary['同系列品类价差比'] = product_bom_summary.apply(
    lambda x: (x[f'最低有效价-wk{current_week}'] - x['同系列品类最低价']) / x['同系列品类最低价'] 
    if x[f'最低有效价-wk{current_week}'] > 0 and x['同系列品类最低价'] > 0 
    else 0,
    axis=1
)

# 计算同项目品类价差比
print('正在计算同项目品类价差比...')
product_bom_summary['同项目品类价差比'] = product_bom_summary.apply(
    lambda x: (x[f'最低有效价-wk{current_week}'] - x['同项目品类最低价']) / x['同项目品类最低价'] 
    if x[f'最低有效价-wk{current_week}'] > 0 and x['同项目品类最低价'] > 0 
    else 0,
    axis=1
)

# 计算涨跌幅
print('正在计算涨跌幅...')
product_bom_summary['涨跌幅'] = product_bom_summary.apply(
    lambda x: (x[f'最低有效价-wk{current_week}'] - x['上月单价']) / x['上月单价']
    if x[f'最低有效价-wk{current_week}'] > 0 and x['上月单价'] > 0
    else 0,
    axis=1
)

# 计算涨价总金额
print('正在计算涨价总金额...')
product_bom_summary['涨价总金额'] = product_bom_summary.apply(
    lambda x: x['涨跌幅'] * x['上月单价'] * x[f'{current_month}月物料预测量']
    if x['涨跌幅'] > 0 and x['上月单价'] > 0 and x[f'{current_month}月物料预测量'] > 0
    else 0,
    axis=1
)

# 计算目标差额
print('正在计算目标差额...')
product_bom_summary['目标差额'] = product_bom_summary.apply(
    lambda x: (x[f'bom成本-wk{current_week}'] - x[f'{current_month}月目标bom成本']) * x[f'{current_month}月整机预测量']
    if x[f'bom成本-wk{current_week}'] > 0 and x[f'{current_month}月目标bom成本'] > 0 and x[f'{current_month}月整机预测量'] > 0
    else 0,
    axis=1
)

# 计算物料目标差额汇总
print('正在计算物料目标差额汇总...')
material_target_diff = product_bom_summary.groupby('物料编码')['目标差额'].sum().reset_index()
material_target_diff = material_target_diff.rename(columns={'目标差额': '物料目标差额汇总'})

# 将物料目标差额汇总连接到主表
product_bom_summary = pd.merge(product_bom_summary, material_target_diff, on='物料编码', how='left')


#%%
columns_order = ['整机编码', '物料描述', '系列', '项目号', 'PC', '产品', '版本', '渠道', '品牌', '工厂', '状态',
    '物料编码', '中文名称', '单位', '中分类', '小分类', '一级分类', '采购', 'BOM用量', '上月单价',
    '价格类型', '供应商', '上月成本', '最低基价(含还原)', '基准bom成本', '财务核价',
    '最低有效价-wk1', '最低有效价-wk2', '最低有效价-wk3', '最低有效价-wk4',
    '1月目标价', '2月目标价', '3月目标价', '4月目标价', '5月目标价', '6月目标价',
    '最终价(目标价&财务核价取低)',
    'bom成本-wk1', 'bom成本-wk2', 'bom成本-wk3', 'bom成本-wk4',
    '1月目标bom成本', '2月目标bom成本', '3月目标bom成本', '4月目标bom成本', '5月目标bom成本', '6月目标bom成本',
    '最终成本(目标价&财务核价取低)',
    '1月整机预测量', '2月整机预测量', '3月整机预测量', '4月整机预测量', '5月整机预测量', '6月整机预测量',
    '7月整机预测量', '8月整机预测量', '9月整机预测量', '10月整机预测量', '11月整机预测量', '12月整机预测量',
    '1月物料预测量', '2月物料预测量', '3月物料预测量', '4月物料预测量', '5月物料预测量', '6月物料预测量',
    '7月物料预测量', '8月物料预测量', '9月物料预测量', '10月物料预测量', '11月物料预测量', '12月物料预测量',
    '系列+小分类', '同系列品类最低价', '同系列品类价差比',
    '项目+小分类', '同项目品类最低价', '同项目品类价差比',
    '涨跌幅', '涨价总金额', '目标差额', '物料目标差额汇总']

product_bom_summary = product_bom_summary[columns_order]
#%%
print_section('物料降本(多供方)')

print('正在读取最低基准价...')
min_price = pd.read_excel(r'input\大表基准价.xlsx')


dup_material = product_bom_summary[['项目号', '物料编码', '中文名称', '单位', '中分类', '小分类', '一级分类','采购',
    '1月物料预测量', '2月物料预测量', '3月物料预测量', '4月物料预测量', '5月物料预测量', '6月物料预测量',
    '7月物料预测量', '8月物料预测量', '9月物料预测量', '10月物料预测量', '11月物料预测量', '12月物料预测量']]
dup_material = dup_material.rename(columns={'一级分类':'组别'})
# 对物料编码进行分组聚合
dup_material = dup_material.groupby('物料编码').agg({
    '中文名称': 'first',
    '单位': 'first',
    '中分类': 'first',
    '小分类': 'first',
    '组别': 'first',
    '采购': 'first',
    '项目号': lambda x: ','.join(x.unique()),
    '1月物料预测量': 'sum',
    '2月物料预测量': 'sum',
    '3月物料预测量': 'sum',
    '4月物料预测量': 'sum',
    '5月物料预测量': 'sum',
    '6月物料预测量': 'sum',
    '7月物料预测量': 'sum',
    '8月物料预测量': 'sum',
    '9月物料预测量': 'sum',
    '10月物料预测量': 'sum',
    '11月物料预测量': 'sum',
    '12月物料预测量': 'sum'
}).reset_index()

#%%
# 读取各周价格

print('正在读取current_price-wk1...')
current_price_wk1 = pd.read_excel(r'input\大表价格-wk1.xlsx')
current_price_wk1 = current_price_wk1.rename(columns={'SAP物料编码':'物料编码'})
# 处理current_price_wk1的筛选逻辑
print('正在处理current_price-wk1的筛选...')
# 重置价格
current_price_wk1 = current_price_wk1.drop(columns=['价格']).rename(columns={'最终价格(含税人民币)':'现价-wk1'})
# 只保留供应商编码长度不为4位的，或者供应商编码=1200的
current_price_wk1 = current_price_wk1[(current_price_wk1['供应商编码'].astype(str).str.len() != 4) | (current_price_wk1['供应商编码'].astype(str) == '1200')]
# 1. 筛选价格在有效期内的记录
current_price_wk1['有效期开始时间'] = pd.to_datetime(current_price_wk1['有效期开始时间'], errors='coerce').dt.date
# 对于9999-12-31的日期，先转换为字符串再处理
current_price_wk1['有效期到期时间'] = current_price_wk1['有效期到期时间'].astype(str)
current_price_wk1.loc[current_price_wk1['有效期到期时间'] == '9999-12-31', '有效期到期时间'] = pd.Timestamp('9999-12-31').date()
current_price_wk1.loc[current_price_wk1['有效期到期时间'] != '9999-12-31', '有效期到期时间'] = pd.to_datetime(current_price_wk1['有效期到期时间'], errors='coerce').dt.date

# 特殊处理9999-12-31的日期
valid_price = current_price_wk1[
    ((current_price_wk1['有效期开始时间'] <= today) & 
    ((current_price_wk1['有效期到期时间'] >= today) | 
     (current_price_wk1['有效期到期时间'] == pd.Timestamp('9999-12-31').date())))
].copy()

# 2. 按物料编码分组，优先取正式价格，其次取暂估价，都取创建日期最新的
valid_price['创建日期'] = pd.to_datetime(valid_price['创建日期'], errors='coerce')
valid_price = valid_price.sort_values('创建日期', ascending=False).drop_duplicates(subset=['物料编码','价格类型','供应商编码'],keep='first')  # 按创建日期降序排序

# 先尝试获取正式价格
formal_price = valid_price[valid_price['价格类型'] == '正式价'].groupby(['物料编码','供应商编码']).first().reset_index()
# 对于没有正式价格的物料，获取暂估价
temp_price = valid_price[valid_price['价格类型'] == '试产价'].groupby(['物料编码','供应商编码']).first().reset_index()
# 合并两种价格
final_price = pd.concat([formal_price, temp_price])
# 只保留需要的列
current_price_wk1 = final_price[['物料编码', '现价-wk1','价格类型','有效期开始时间','有效期到期时间','供应商编码','供应商描述','系统配额','修正配额']].sort_values(by='现价-wk1', ascending=True).drop_duplicates(subset=['物料编码','供应商编码'],keep='first')

print('正在读取current_price-wk2...')
current_price_wk2 = pd.read_excel(r'input\大表价格-wk2.xlsx')
current_price_wk2 = current_price_wk2.rename(columns={'SAP物料编码':'物料编码'})
# 处理current_price_wk1的筛选逻辑
print('正在处理current_price-wk2的筛选...')
# 重置价格
current_price_wk2 = current_price_wk2.drop(columns=['价格']).rename(columns={'最终价格(含税人民币)':'现价-wk2'})
# 只保留供应商编码长度不为4位的，或者供应商编码=1200的
current_price_wk2 = current_price_wk2[(current_price_wk2['供应商编码'].astype(str).str.len() != 4) | (current_price_wk2['供应商编码'].astype(str) == '1200')]
# 1. 筛选价格在有效期内的记录
current_price_wk2['有效期开始时间'] = pd.to_datetime(current_price_wk2['有效期开始时间'], errors='coerce').dt.date
# 对于9999-12-31的日期，先转换为字符串再处理
current_price_wk2['有效期到期时间'] = current_price_wk2['有效期到期时间'].astype(str)
current_price_wk2.loc[current_price_wk2['有效期到期时间'] == '9999-12-31', '有效期到期时间'] = pd.Timestamp('9999-12-31').date()
current_price_wk2.loc[current_price_wk2['有效期到期时间'] != '9999-12-31', '有效期到期时间'] = pd.to_datetime(current_price_wk2['有效期到期时间'], errors='coerce').dt.date

# 特殊处理9999-12-31的日期
valid_price = current_price_wk2[
    ((current_price_wk2['有效期开始时间'] <= today) & 
    ((current_price_wk2['有效期到期时间'] >= today) | 
     (current_price_wk2['有效期到期时间'] == pd.Timestamp('9999-12-31').date())))
].copy()

# 2. 按物料编码分组，优先取正式价格，其次取暂估价，都取创建日期最新的
valid_price['创建日期'] = pd.to_datetime(valid_price['创建日期'], errors='coerce')
valid_price = valid_price.sort_values('创建日期', ascending=False).drop_duplicates(subset=['物料编码','价格类型','供应商编码'],keep='first')  # 按创建日期降序排序

# 先尝试获取正式价格
formal_price = valid_price[valid_price['价格类型'] == '正式价'].groupby(['物料编码','供应商编码']).first().reset_index()
# 对于没有正式价格的物料，获取暂估价
temp_price = valid_price[valid_price['价格类型'] == '试产价'].groupby(['物料编码','供应商编码']).first().reset_index()
# 合并两种价格
final_price = pd.concat([formal_price, temp_price])
# 重命名价格列
final_price = final_price.rename(columns={'价格': '现价-wk2'})
# 只保留需要的列
current_price_wk2 = final_price[['物料编码', '现价-wk2','价格类型','有效期开始时间','有效期到期时间','供应商编码','供应商描述','系统配额','修正配额']].sort_values(by='现价-wk2', ascending=True).drop_duplicates(subset=['物料编码','供应商编码'],keep='first')

print('正在读取current_price-wk3...')
current_price_wk3 = pd.read_excel(r'input\大表价格-wk3.xlsx')
current_price_wk3 = current_price_wk3.rename(columns={'SAP物料编码':'物料编码'})
# 处理current_price_wk1的筛选逻辑
print('正在处理current_price-wk3的筛选...')
# 重置价格
current_price_wk3 = current_price_wk3.drop(columns=['价格']).rename(columns={'最终价格(含税人民币)':'现价-wk3'})
# 只保留供应商编码长度不为4位的，或者供应商编码=1200的
current_price_wk3 = current_price_wk3[(current_price_wk3['供应商编码'].astype(str).str.len() != 4) | (current_price_wk3['供应商编码'].astype(str) == '1200')]
# 1. 筛选价格在有效期内的记录
current_price_wk3['有效期开始时间'] = pd.to_datetime(current_price_wk3['有效期开始时间'], errors='coerce').dt.date
# 对于9999-12-31的日期，先转换为字符串再处理
current_price_wk3['有效期到期时间'] = current_price_wk3['有效期到期时间'].astype(str)
current_price_wk3.loc[current_price_wk3['有效期到期时间'] == '9999-12-31', '有效期到期时间'] = pd.Timestamp('9999-12-31').date()
current_price_wk3.loc[current_price_wk3['有效期到期时间'] != '9999-12-31', '有效期到期时间'] = pd.to_datetime(current_price_wk3['有效期到期时间'], errors='coerce').dt.date

# 特殊处理9999-12-31的日期
valid_price = current_price_wk3[
    ((current_price_wk3['有效期开始时间'] <= today) & 
    ((current_price_wk3['有效期到期时间'] >= today) | 
     (current_price_wk3['有效期到期时间'] == pd.Timestamp('9999-12-31').date())))
].copy()

# 2. 按物料编码分组，优先取正式价格，其次取暂估价，都取创建日期最新的
valid_price['创建日期'] = pd.to_datetime(valid_price['创建日期'], errors='coerce')
valid_price = valid_price.sort_values('创建日期', ascending=False).drop_duplicates(subset=['物料编码','价格类型','供应商编码'],keep='first')  # 按创建日期降序排序

# 先尝试获取正式价格
formal_price = valid_price[valid_price['价格类型'] == '正式价'].groupby(['物料编码','供应商编码']).first().reset_index()
# 对于没有正式价格的物料，获取暂估价
temp_price = valid_price[valid_price['价格类型'] == '试产价'].groupby(['物料编码','供应商编码']).first().reset_index()

# 合并两种价格
final_price = pd.concat([formal_price, temp_price])
# 重命名价格列
final_price = final_price.rename(columns={'价格': '现价-wk3'})
# 只保留需要的列
current_price_wk3 = final_price[['物料编码', '现价-wk3','价格类型','有效期开始时间','有效期到期时间','供应商编码','供应商描述','系统配额','修正配额']].sort_values(by='现价-wk3', ascending=True).drop_duplicates(subset=['物料编码','供应商编码'],keep='first')

print('正在读取current_price-wk4...')
current_price_wk4 = pd.read_excel(r'input\大表价格-wk4.xlsx')
current_price_wk4 = current_price_wk4.rename(columns={'SAP物料编码':'物料编码'})
# 处理current_price_wk1的筛选逻辑
print('正在处理current_price-wk4的筛选...')
# 重置价格
current_price_wk4 = current_price_wk4.drop(columns=['价格']).rename(columns={'最终价格(含税人民币)':'现价-wk4'})
# 只保留供应商编码长度不为4位的，或者供应商编码=1200的
current_price_wk4 = current_price_wk4[(current_price_wk4['供应商编码'].astype(str).str.len() != 4) | (current_price_wk4['供应商编码'].astype(str) == '1200')]
# 1. 筛选价格在有效期内的记录
current_price_wk4['有效期开始时间'] = pd.to_datetime(current_price_wk4['有效期开始时间'], errors='coerce').dt.date
# 对于9999-12-31的日期，先转换为字符串再处理
current_price_wk4['有效期到期时间'] = current_price_wk4['有效期到期时间'].astype(str)
current_price_wk4.loc[current_price_wk4['有效期到期时间'] == '9999-12-31', '有效期到期时间'] = pd.Timestamp('9999-12-31').date()
current_price_wk4.loc[current_price_wk4['有效期到期时间'] != '9999-12-31', '有效期到期时间'] = pd.to_datetime(current_price_wk4['有效期到期时间'], errors='coerce').dt.date

# 特殊处理9999-12-31的日期
valid_price = current_price_wk4[
    ((current_price_wk4['有效期开始时间'] <= today) & 
    ((current_price_wk4['有效期到期时间'] >= today) | 
     (current_price_wk4['有效期到期时间'] == pd.Timestamp('9999-12-31').date())))
].copy()

# 2. 按物料编码分组，优先取正式价格，其次取暂估价，都取创建日期最新的
valid_price['创建日期'] = pd.to_datetime(valid_price['创建日期'], errors='coerce')
valid_price = valid_price.sort_values('创建日期', ascending=False).drop_duplicates(subset=['物料编码','价格类型','供应商编码'],keep='first')  # 按创建日期降序排序

# 先尝试获取正式价格
formal_price = valid_price[valid_price['价格类型'] == '正式价'].groupby(['物料编码','供应商编码']).first().reset_index()
# 对于没有正式价格的物料，获取暂估价
temp_price = valid_price[valid_price['价格类型'] == '试产价'].groupby(['物料编码','供应商编码']).first().reset_index()
# 合并两种价格
final_price = pd.concat([formal_price, temp_price])
# 只保留需要的列
current_price_wk4 = final_price[['物料编码', '现价-wk4','价格类型','有效期开始时间','有效期到期时间','供应商编码','供应商描述','系统配额','修正配额']].sort_values(by='现价-wk4', ascending=True).drop_duplicates(subset=['物料编码','供应商编码'],keep='first')
#%%
print('正在根据current_week处理current_price_wkn...')
weeks_to_process = {1: [2, 3, 4], 2: [1, 3, 4], 3: [1, 2, 4], 4: [1, 2, 3]}
weeks = weeks_to_process[current_week]

for week in weeks:
    df_name = f'current_price_wk{week}'
    price_col = f'现价-wk{week}'
    df = locals()[df_name]
    df = df.drop(columns=['价格类型','有效期开始时间','有效期到期时间','系统配额','修正配额']).groupby(['物料编码', '供应商编码']).agg({
        price_col: 'first',
    }).reset_index()
    locals()[df_name] = df

print('正在根据current_week进行DataFrame连接...')

# 首先与当前周的current_price_wkn进行连接
current_week_df = locals()[f'current_price_wk{current_week}']
result = pd.merge(dup_material, current_week_df, on='物料编码', how='left')

# 然后与其他周的current_price_wk进行连接
for week in weeks:
    df_name = f'current_price_wk{week}'
    df = locals()[df_name]
    result = pd.merge(result, df, on=['物料编码', '供应商编码'], how='left')

# 将大于当前周的现价设置为当前周的现价
print(f'正在将大于当前周(wk{current_week})的现价设置为当前周的现价...')
current_week_price_col = f'现价-wk{current_week}'

# 获取大于当前周的所有周
future_weeks = [i for i in range(1, 5) if i > current_week]

for week in future_weeks:
    future_week_price_col = f'现价-wk{week}'
    # 将大于当前周的现价设置为当前周的现价
    result[future_week_price_col] = result[current_week_price_col]
    print(f'已将{future_week_price_col}设置为{current_week_price_col}')


print('DataFrame连接完成！')
#%%
print('修正各月预测量')
# 确保修正配额是数值类型并处理缺失值
result['修正配额'] = pd.to_numeric(result['修正配额'], errors='coerce').fillna(0)

# 对每个月的物料预测量进行修正
for month in range(1, 13):
    month_col = f'{month}月物料预测量'
    if month_col in result.columns:
        # 将物料预测量乘以修正配额
        result[month_col] = result.apply(
            lambda x: x[month_col] * x['修正配额']/100 if pd.notna(x[month_col]) and x['修正配额'] >= 0 else 0,
            axis=1
        )

print('各月物料预测量修正完成！')
#%%
#%%
print('读取财务核价物料清单.xlsx')
financial_price = pd.read_excel(r'input\财务核价物料清单.xlsx')
financial_price = financial_price.rename(columns={'SAP物料编码':'物料编码'})[['物料编码','财务核价']].drop_duplicates(subset=['物料编码'],keep='first')
print('读取加权基价')
avg_price = pd.read_excel(r'input\大表基准价.xlsx')
avg_price = avg_price[['SAP物料编码','加权基价']].drop_duplicates(subset=['SAP物料编码'])
avg_price.rename(columns={'SAP物料编码':'物料编码','加权基价':'加权基价'},inplace=True)

print('正在连接财务核价和加权基价到dup_material...')
# 将财务核价连接到result
result = pd.merge(result, financial_price, on='物料编码', how='left')
# 将加权基价连接到result
result = pd.merge(result, avg_price, on='物料编码', how='left')
print('财务核价和加权基价连接完成！')

print('正在读取采购目标价')
purchase_target_price = pd.read_excel(r'input\采购目标价.xlsx')
purchase_target_price = purchase_target_price[['物料编码','供应商描述','W1目标价','W2目标价','W3目标价','W4目标价','1月目标价','2月目标价','3月目标价','4月目标价','5月目标价','6月目标价','7月目标价','8月目标价','9月目标价','10月目标价','11月目标价','12月目标价']].drop_duplicates(subset=['物料编码','供应商描述'],keep='first')

print('正在连接采购目标价到result...')
result = pd.merge(result, purchase_target_price, on=['物料编码','供应商描述'], how='left')
print('采购目标价连接完成！')
#%%
print('正在计算降本率和GAP...')

# 获取当前周的现价列名
current_week_price_col = f'现价-wk{current_week}'
# 获取当前月的目标价列名
current_month_target_col = f'{current_month}月目标价'

# 计算降本率(基价vs现价)
result['降本率(基价vs现价)'] = result.apply(
    lambda x: (x['加权基价'] - x[current_week_price_col]) / x['加权基价'] 
    if pd.notna(x['加权基价']) and pd.notna(x[current_week_price_col]) 
    and x['加权基价'] != 0 and x[current_week_price_col] != 0 
    else 0,
    axis=1
)

# 计算GAP(现价-目标)
result['GAP(现价-目标)'] = result.apply(
    lambda x: x[current_week_price_col] - x[current_month_target_col]
    if pd.notna(x[current_week_price_col]) and pd.notna(x[current_month_target_col])
    and x[current_week_price_col] != 0 and x[current_month_target_col] != 0
    else 0,
    axis=1
)

print('正在计算GAP降本额...')
# 获取当前月的物料预测量列名
current_month_forecast_col = f'{current_month}月物料预测量'

# 计算GAP降本额/万
result['GAP降本额/万'] = result.apply(
    lambda x: x['GAP(现价-目标)'] * x[current_month_forecast_col] / 10000
    if pd.notna(x['GAP(现价-目标)']) and pd.notna(x[current_month_forecast_col])
    else 0,
    axis=1
)
#%%
print('正在计算各周降本额...')
# 计算各周的降本额
for week in range(1, 5):
    price_col = f'现价-wk{week}'
    result[f'降本额-wk{week}'] = result.apply(
        lambda x: (x['加权基价'] - x[price_col]) * x[current_month_forecast_col] / 10000
        if pd.notna(x['加权基价']) and pd.notna(x[price_col]) 
        and x['加权基价'] != 0 and x[price_col] != 0
        else 0,
        axis=1
    )

print('正在计算各月相关金额...')
# 计算各月的相关金额
for month in range(1, 13):
    target_col = f'{month}月目标价'
    forecast_col = f'{month}月物料预测量'
    
    # 计算预测采购额
    result[f'{month}月预测采购额'] = result.apply(
        lambda x: x[target_col] * x[forecast_col] / 10000
        if pd.notna(x[target_col]) and x[target_col] != 0
        else x['现价-wk4'] * x[forecast_col] / 10000
        if pd.notna(x['现价-wk4']) and x['现价-wk4'] != 0
        else 0,
        axis=1
    )
    
    # 计算基准金额
    result[f'{month}月基准金额'] = result.apply(
        lambda x: x['加权基价'] * x[forecast_col] / 10000
        if pd.notna(x['加权基价']) and x['加权基价'] != 0
        else x['现价-wk4'] * x[forecast_col] / 10000
        if pd.notna(x['现价-wk4']) and x['现价-wk4'] != 0
        else 0,
        axis=1
    )
    
    # 计算预测降本额
    result[f'{month}月预测降本额'] = result.apply(
        lambda x: (x['加权基价'] - x[target_col]) * x[forecast_col] / 10000
        if pd.notna(x['加权基价']) and pd.notna(x[target_col]) 
        and x['加权基价'] != 0 and x[target_col] != 0
        else (x['加权基价'] - x['现价-wk4']) * x[forecast_col] / 10000
        if pd.notna(x['加权基价']) and pd.notna(x['现价-wk4'])
        and x['加权基价'] != 0 and x['现价-wk4'] != 0
        else 0,
        axis=1
    )

print('所有计算列添加完成！')
#%%

# %%

print('正在调整列顺序...')

# 定义新的列顺序
new_columns = [
    '中分类', '小分类', '组别', '采购', '物料编码', '中文名称', '单位', '项目号', '财务核价', '加权基价',
    '降本率(基价vs现价)', 'GAP(现价-目标)', 'GAP降本额/万', '现价-wk1', '现价-wk2', '现价-wk3', '现价-wk4',
    '降本额-wk1', '降本额-wk2', '降本额-wk3', '降本额-wk4', '价格类型', '有效期开始时间', '有效期到期时间',
    '供应商描述', '系统配额', '修正配额', 'W1目标价', 'W2目标价', 'W3目标价', 'W4目标价', '1月目标价',
    '2月目标价', '3月目标价', '4月目标价', '5月目标价', '6月目标价', '7月目标价', '8月目标价',
    '9月目标价', '10月目标价', '11月目标价', '12月目标价', '1月物料预测量', '2月物料预测量', '3月物料预测量',
    '4月物料预测量', '5月物料预测量', '6月物料预测量', '7月物料预测量', '8月物料预测量', '9月物料预测量',
    '10月物料预测量', '11月物料预测量', '12月物料预测量', '1月预测采购额', '2月预测采购额', '3月预测采购额',
    '4月预测采购额', '5月预测采购额', '6月预测采购额', '7月预测采购额', '8月预测采购额', '9月预测采购额',
    '10月预测采购额', '11月预测采购额', '12月预测采购额', '1月基准金额', '2月基准金额', '3月基准金额',
    '4月基准金额', '5月基准金额', '6月基准金额', '7月基准金额', '8月基准金额', '9月基准金额',
    '10月基准金额', '11月基准金额', '12月基准金额', '1月预测降本额', '2月预测降本额', '3月预测降本额',
    '4月预测降本额', '5月预测降本额', '6月预测降本额', '7月预测降本额', '8月预测降本额', '9月预测降本额',
    '10月预测降本额', '11月预测降本额', '12月预测降本额'
]

# 添加缺失的列（如果不存在）
for col in new_columns:
    if col not in result.columns:
        result[col] = None

# 重新排序列
result = result[new_columns]
print('列顺序调整完成！')

print('正在生成文件...')
# 生成时间后缀
current_time = datetime.now()
time_suffix = current_time.strftime("%Y%m%d_%H%M")
print('正在生成机型-F物料文件...')
product_bom_summary.to_excel(r'output\机型-F物料{}.xlsx'.format(time_suffix),index=False)
print('机型-F物料文件生成完成！')
print('正在生成物料降本-多供方文件...')
result.to_excel(r'output\物料降本-多供方{}.xlsx'.format(time_suffix),index=False)
print('物料降本-多供方文件生成完成！')

print('将机型-F物料文件处理后并删除时间后缀，保存到processing文件夹，以运行下一步')
input("""

            ______ _   __ ____              
           / ____// | / // __ \             
 ______   / __/  /  |/ // / / /  ______     
/_____/  / /___ / /|  // /_/ /  /_____/     
        /_____//_/ |_//_____/               
                                            

"""
)

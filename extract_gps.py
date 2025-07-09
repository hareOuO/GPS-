import pandas as pd
import re

# 读取Excel文件
df = pd.read_excel('1.xlsx')

# 检查是否存在"事件地点"列
required_columns = ['事件地点']
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    print(f"错误：Excel文件中缺少以下列：{', '.join(missing_columns)}")
    exit(1)

# 检查是否存在"车辆VIN"列
has_vin = '车辆VIN' in df.columns
if has_vin:
    print("发现'车辆VIN'列，将提取VIN码")
else:
    print("未找到'车辆VIN'列，将使用空值代替")

# 检查是否存在"事件时间"列
has_time = '事件时间' in df.columns
if has_time:
    print("发现'事件时间'列，将提取事件时间")
else:
    print("未找到'事件时间'列，将使用空值代替")

# 提取GPS坐标的函数
def extract_gps(location_str):
    if not isinstance(location_str, str):
        return None, None
    
    # 匹配GPS坐标的正则表达式 (匹配两个浮点数，用逗号分隔)
    pattern = r'(\d+\.\d+)\s*,\s*(\d+\.\d+)'
    match = re.search(pattern, location_str)
    
    if match:
        longitude = float(match.group(1))
        latitude = float(match.group(2))
        return longitude, latitude
    else:
        return None, None

# 提取坐标
results = []
for idx, row in df.iterrows():
    location = row['事件地点']
    lon, lat = extract_gps(location)
    if lon is not None and lat is not None:
        result = {'longitude': lon, 'latitude': lat}
        
        # 如果存在VIN列，则提取VIN码
        if has_vin:
            vin = row['车辆VIN'] if pd.notna(row['车辆VIN']) else "未知VIN"
            result['vin'] = str(vin)
        else:
            result['vin'] = "未知VIN"
        
        # 如果存在事件时间列，则提取事件时间
        if has_time:
            event_time = row['事件时间'] if pd.notna(row['事件时间']) else "未知时间"
            # 如果事件时间是datetime类型，转换为字符串
            if pd.api.types.is_datetime64_any_dtype(event_time):
                event_time = event_time.strftime('%Y-%m-%d %H:%M:%S')
            result['event_time'] = str(event_time)
        else:
            result['event_time'] = "未知时间"
            
        results.append(result)

# 创建新的DataFrame并保存
if results:
    result_df = pd.DataFrame(results)
    result_df.to_csv('gps_coordinates.csv', index=False)
    result_df.to_excel('gps_coordinates.xlsx', index=False)
    print(f"成功提取了 {len(results)} 条GPS坐标")
    print(f"数据已保存到 gps_coordinates.csv 和 gps_coordinates.xlsx")
else:
    print("没有找到有效的GPS坐标") 
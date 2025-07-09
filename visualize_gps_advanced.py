import pandas as pd
import folium
from folium.plugins import HeatMap, MarkerCluster
import os
import webbrowser

print("开始执行高级地图可视化脚本...")

# 检查文件是否存在
gps_file = 'gps_coordinates.xlsx'
if not os.path.exists(gps_file):
    # 如果提取的文件不存在，先运行提取脚本
    print(f"文件 {gps_file} 不存在，尝试从原始数据提取...")
    
    import re
    try:
        df_original = pd.read_excel('1.xlsx')
        print("成功读取原始Excel文件")
        
        # 检查是否存在"事件地点"列
        if '事件地点' not in df_original.columns:
            print("错误：Excel文件中没有'事件地点'列")
            exit(1)
        
        # 检查是否存在"车辆VIN"列
        has_vin = '车辆VIN' in df_original.columns
        if has_vin:
            print("发现'车辆VIN'列，将提取VIN码")
        else:
            print("未找到'车辆VIN'列，将使用空值代替")
            
        # 检查是否存在"事件时间"列
        has_time = '事件时间' in df_original.columns
        if has_time:
            print("发现'事件时间'列，将提取事件时间")
        else:
            print("未找到'事件时间'列，将使用空值代替")
        
        # 提取GPS坐标的函数
        def extract_gps(location_str):
            if not isinstance(location_str, str):
                return None, None
            
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
        for idx, row in df_original.iterrows():
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
        
        if not results:
            print("没有找到有效的GPS坐标")
            exit(1)
            
        df = pd.DataFrame(results)
        
        # 保存提取的坐标
        df.to_excel(gps_file, index=False)
        print(f"已将提取的坐标保存到 {gps_file}")
    except Exception as e:
        print(f"提取坐标时出错: {str(e)}")
        exit(1)
else:
    # 读取已提取的坐标
    try:
        df = pd.read_excel(gps_file)
        print(f"成功读取 {gps_file}")
    except Exception as e:
        print(f"读取文件 {gps_file} 时出错: {str(e)}")
        exit(1)

# 打印数据的前几行，确认数据格式正确
print("\n数据前5行:")
print(df.head())

# 确保数据中有经纬度列
if 'longitude' not in df.columns or 'latitude' not in df.columns:
    print("错误：数据中缺少经度(longitude)或纬度(latitude)列")
    exit(1)

# 检查是否有VIN列
has_vin = 'vin' in df.columns
if not has_vin:
    print("未找到VIN码列，将使用默认标签")
    df['vin'] = "未知VIN"

# 检查是否有事件时间列
has_time = 'event_time' in df.columns
if not has_time:
    print("未找到事件时间列，将使用默认标签")
    df['event_time'] = "未知时间"

# 检查坐标是否在合理范围内
china_lon_min, china_lon_max = 73, 135
china_lat_min, china_lat_max = 18, 53

# 筛选有效点
valid_points = (df['longitude'] >= china_lon_min) & (df['longitude'] <= china_lon_max) & \
               (df['latitude'] >= china_lat_min) & (df['latitude'] <= china_lat_max)

valid_count = sum(valid_points)
invalid_count = len(df) - valid_count
print(f"\n有效坐标点: {valid_count}，异常坐标点: {invalid_count}")

df_valid = df[valid_points]
if invalid_count > 0:
    df_invalid = df[~valid_points]

# 确保有有效的坐标点
if valid_count == 0:
    print("错误：没有有效的坐标点")
    exit(1)

# 计算中心点
center_lat = df_valid['latitude'].mean()
center_lon = df_valid['longitude'].mean()
print(f"地图中心点: 经度={center_lon}, 纬度={center_lat}")

# 创建多个地图，以不同方式展示数据
# 1. 基础地图 - 使用圆形标记
base_map = folium.Map(location=[center_lat, center_lon], zoom_start=5, tiles='CartoDB positron')

# 2. 热力图
heat_map = folium.Map(location=[center_lat, center_lon], zoom_start=5, tiles='CartoDB positron')

# 3. 聚类地图
cluster_map = folium.Map(location=[center_lat, center_lon], zoom_start=5, tiles='CartoDB positron')

# 添加标记到基础地图
print("添加标记到基础地图...")
for idx, row in df_valid.iterrows():
    # 准备弹出窗口内容，包含VIN码和事件时间
    popup_content = f"""
    <b>VIN码:</b> {row['vin']}<br>
    <b>事件时间:</b> {row['event_time']}<br>
    <b>坐标:</b> {row['longitude']}, {row['latitude']}<br>
    <b>类型:</b> 有效坐标
    """
    
    folium.CircleMarker(
        location=[row['latitude'], row['longitude']],
        radius=8,  # 更大的半径
        color='red',
        fill=True,
        fill_color='red',
        fill_opacity=0.7,
        popup=folium.Popup(popup_content, max_width=300)
    ).add_to(base_map)

# 如果有异常点，也添加到基础地图
if invalid_count > 0:
    for idx, row in df_invalid.iterrows():
        # 准备弹出窗口内容，包含VIN码和事件时间
        popup_content = f"""
        <b>VIN码:</b> {row['vin']}<br>
        <b>事件时间:</b> {row['event_time']}<br>
        <b>坐标:</b> {row['longitude']}, {row['latitude']}<br>
        <b>类型:</b> 异常坐标
        """
        
        folium.CircleMarker(
            location=[row['latitude'], row['longitude']],
            radius=8,
            color='blue',
            fill=True,
            fill_color='blue',
            fill_opacity=0.7,
            popup=folium.Popup(popup_content, max_width=300)
        ).add_to(base_map)

# 添加热力图
print("创建热力图...")
# 将所有点（包括有效点和异常点）都添加到热力图
heat_data = [[row['latitude'], row['longitude']] for _, row in df.iterrows()]
HeatMap(heat_data, radius=15, gradient={0.4: 'blue', 0.65: 'lime', 0.8: 'orange', 1: 'red'}).add_to(heat_map)

# 添加聚类标记
print("创建聚类地图...")
# 为有效点创建聚类
valid_marker_cluster = MarkerCluster(name="有效坐标聚类").add_to(cluster_map)
for idx, row in df_valid.iterrows():
    # 准备弹出窗口内容，包含VIN码和事件时间
    popup_content = f"""
    <b>VIN码:</b> {row['vin']}<br>
    <b>事件时间:</b> {row['event_time']}<br>
    <b>坐标:</b> {row['longitude']}, {row['latitude']}<br>
    <b>类型:</b> 有效坐标
    """
    
    folium.Marker(
        location=[row['latitude'], row['longitude']],
        popup=folium.Popup(popup_content, max_width=300),
        icon=folium.Icon(color='red', icon='info-sign')
    ).add_to(valid_marker_cluster)

# 为异常点创建单独的聚类
if invalid_count > 0:
    print("添加异常点到聚类图...")
    invalid_marker_cluster = MarkerCluster(name="异常坐标聚类").add_to(cluster_map)
    for idx, row in df_invalid.iterrows():
        # 准备弹出窗口内容，包含VIN码和事件时间
        popup_content = f"""
        <b>VIN码:</b> {row['vin']}<br>
        <b>事件时间:</b> {row['event_time']}<br>
        <b>坐标:</b> {row['longitude']}, {row['latitude']}<br>
        <b>类型:</b> 异常坐标
        """
        
        folium.Marker(
            location=[row['latitude'], row['longitude']],
            popup=folium.Popup(popup_content, max_width=300),
            icon=folium.Icon(color='blue', icon='info-sign')
        ).add_to(invalid_marker_cluster)

# 添加图层控制到聚类地图
folium.LayerControl().add_to(cluster_map)

# 保存所有地图
base_map.save('base_map.html')
heat_map.save('heat_map.html')
cluster_map.save('cluster_map.html')

print(f"已生成三个地图文件:")
print(f"1. 基础地图: {os.path.abspath('base_map.html')}")
print(f"2. 热力图: {os.path.abspath('heat_map.html')}")
print(f"3. 聚类地图: {os.path.abspath('cluster_map.html')}")

# 尝试自动打开基础地图
# try:
#     print("尝试在浏览器中打开基础地图...")
#     webbrowser.open(os.path.abspath('base_map.html'))
# except Exception as e:
#     print(f"无法自动打开地图: {str(e)}")
#     print("请手动打开生成的HTML文件查看地图") 
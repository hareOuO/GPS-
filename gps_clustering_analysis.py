import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans, DBSCAN
from sklearn.preprocessing import StandardScaler
import folium
from folium.plugins import MarkerCluster
import os
import webbrowser
from sklearn.metrics import silhouette_score

print("开始GPS坐标聚类分析...")

# 读取GPS坐标数据
gps_file = 'gps_coordinates.xlsx'
try:
    df = pd.read_excel(gps_file)
    print(f"成功读取 {gps_file}")
except Exception as e:
    print(f"读取文件 {gps_file} 时出错: {str(e)}")
    exit(1)

# 确保数据中有经纬度列
if 'longitude' not in df.columns or 'latitude' not in df.columns:
    print("错误：数据中缺少经度(longitude)或纬度(latitude)列")
    exit(1)

# 检查是否有VIN列
has_vin = 'vin' in df.columns
if has_vin:
    print("发现VIN码列，将在地图上显示VIN信息")
else:
    print("未找到VIN码列，将使用默认标签")
    df['vin'] = "未知VIN"

# 检查是否有事件时间列
has_time = 'event_time' in df.columns
if has_time:
    print("发现事件时间列，将在地图上显示时间信息")
else:
    print("未找到事件时间列，将使用默认标签")
    df['event_time'] = "未知时间"

# 提取所有坐标点（包括异常点）用于聚类分析
X = df[['longitude', 'latitude']].values
print(f"总共有 {len(X)} 个坐标点用于聚类分析")

# 数据标准化（对聚类很重要）
X_scaled = StandardScaler().fit_transform(X)

# 1. K-means聚类分析
print("\n执行K-means聚类分析...")

# 使用肘部法则确定最佳K值
inertia = []
silhouette_scores = []
k_range = range(2, min(11, len(X)))  # 从2到10（或数据点数量）
for k in k_range:
    kmeans = KMeans(n_clusters=k, random_state=42)
    kmeans.fit(X_scaled)
    inertia.append(kmeans.inertia_)
    
    # 计算轮廓系数（只有当k小于数据点数量时才有效）
    if k < len(X):
        score = silhouette_score(X_scaled, kmeans.labels_)
        silhouette_scores.append(score)
        print(f"K={k}, 轮廓系数={score:.4f}")

# 绘制肘部法则图
plt.figure(figsize=(12, 5))
plt.subplot(1, 2, 1)
plt.plot(list(k_range), inertia, marker='o')
plt.xlabel('聚类数量 (K)')
plt.ylabel('惯性')
plt.title('K-means肘部法则')
plt.grid(True)

# 绘制轮廓系数图
plt.subplot(1, 2, 2)
plt.plot(list(k_range), silhouette_scores, marker='o')
plt.xlabel('聚类数量 (K)')
plt.ylabel('轮廓系数')
plt.title('K-means轮廓系数')
plt.grid(True)
plt.tight_layout()
plt.savefig('kmeans_elbow_method.png', dpi=300)
print("已保存K-means肘部法则和轮廓系数图到 kmeans_elbow_method.png")

# 根据轮廓系数选择最佳K值
best_k = k_range[np.argmax(silhouette_scores)]
print(f"根据轮廓系数，最佳聚类数量K = {best_k}")

# 使用最佳K值进行K-means聚类
kmeans = KMeans(n_clusters=best_k, random_state=42)
df['kmeans_cluster'] = kmeans.fit_predict(X_scaled)
print(f"K-means聚类完成，共分为 {best_k} 个类别")

# 2. DBSCAN聚类分析
print("\n执行DBSCAN聚类分析...")
# DBSCAN参数
eps = 0.5  # 邻域半径
min_samples = 5  # 核心点的最小邻居数

dbscan = DBSCAN(eps=eps, min_samples=min_samples)
df['dbscan_cluster'] = dbscan.fit_predict(X_scaled)

# 统计DBSCAN聚类结果
n_clusters = len(set(df['dbscan_cluster'])) - (1 if -1 in df['dbscan_cluster'] else 0)
n_noise = list(df['dbscan_cluster']).count(-1)
print(f"DBSCAN聚类完成，共分为 {n_clusters} 个类别，有 {n_noise} 个噪声点")

# 保存聚类结果
df.to_excel('gps_clustering_results.xlsx', index=False)
print("聚类结果已保存到 gps_clustering_results.xlsx")

# 可视化K-means聚类结果
print("\n创建K-means聚类结果地图...")
kmeans_map = folium.Map(location=[df['latitude'].mean(), df['longitude'].mean()], zoom_start=5)

# 颜色列表，用于不同的聚类
colors = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 
          'lightred', 'beige', 'darkblue', 'darkgreen', 'cadetblue', 
          'darkpurple', 'pink', 'lightblue', 'lightgreen']

# 为每个聚类创建一个图层组
for cluster in range(best_k):
    fg = folium.FeatureGroup(name=f'K-means聚类 {cluster}')
    cluster_points = df[df['kmeans_cluster'] == cluster]
    
    # 确保颜色索引不超出范围
    color_idx = cluster % len(colors)
    
    # 添加该聚类的所有点
    for _, row in cluster_points.iterrows():
        # 准备弹出窗口内容，包含VIN码和事件时间
        popup_content = f"""
        <b>VIN码:</b> {row['vin']}<br>
        <b>事件时间:</b> {row['event_time']}<br>
        <b>坐标:</b> {row['longitude']}, {row['latitude']}<br>
        <b>聚类:</b> {cluster}
        """
        
        folium.CircleMarker(
            location=[row['latitude'], row['longitude']],
            radius=8,
            color=colors[color_idx],
            fill=True,
            fill_color=colors[color_idx],
            fill_opacity=0.7,
            popup=folium.Popup(popup_content, max_width=300)
        ).add_to(fg)
    
    # 添加聚类中心
    cluster_center = kmeans.cluster_centers_[cluster]
    # 需要反标准化聚类中心
    scaler = StandardScaler()
    scaler.fit(X)
    center_original = scaler.inverse_transform([cluster_center])[0]
    
    folium.Marker(
        location=[center_original[1], center_original[0]],
        popup=f"聚类 {cluster} 中心",
        icon=folium.Icon(color=colors[color_idx], icon='info-sign')
    ).add_to(fg)
    
    fg.add_to(kmeans_map)

# 添加图层控制
folium.LayerControl().add_to(kmeans_map)

# 保存K-means聚类地图
kmeans_map.save('kmeans_cluster_map.html')
print("K-means聚类地图已保存到 kmeans_cluster_map.html")

# 可视化DBSCAN聚类结果
print("\n创建DBSCAN聚类结果地图...")
dbscan_map = folium.Map(location=[df['latitude'].mean(), df['longitude'].mean()], zoom_start=5)

# 为每个DBSCAN聚类创建一个图层组
unique_clusters = sorted(df['dbscan_cluster'].unique())
for i, cluster in enumerate(unique_clusters):
    cluster_name = f'噪声点' if cluster == -1 else f'DBSCAN聚类 {cluster}'
    fg = folium.FeatureGroup(name=cluster_name)
    
    cluster_points = df[df['dbscan_cluster'] == cluster]
    
    # 确保颜色索引不超出范围
    color_idx = (i % len(colors)) if cluster != -1 else 0
    point_color = 'gray' if cluster == -1 else colors[color_idx]
    
    # 添加该聚类的所有点
    for _, row in cluster_points.iterrows():
        # 准备弹出窗口内容，包含VIN码和事件时间
        popup_content = f"""
        <b>VIN码:</b> {row['vin']}<br>
        <b>事件时间:</b> {row['event_time']}<br>
        <b>坐标:</b> {row['longitude']}, {row['latitude']}<br>
        <b>聚类:</b> {cluster_name}
        """
        
        folium.CircleMarker(
            location=[row['latitude'], row['longitude']],
            radius=8,
            color=point_color,
            fill=True,
            fill_color=point_color,
            fill_opacity=0.7,
            popup=folium.Popup(popup_content, max_width=300)
        ).add_to(fg)
    
    fg.add_to(dbscan_map)

# 添加图层控制
folium.LayerControl().add_to(dbscan_map)

# 保存DBSCAN聚类地图
dbscan_map.save('dbscan_cluster_map.html')
print("DBSCAN聚类地图已保存到 dbscan_cluster_map.html")

# 尝试自动打开K-means聚类地图
try:
    print("尝试在浏览器中打开K-means聚类地图...")
    webbrowser.open(os.path.abspath('kmeans_cluster_map.html'))
except Exception as e:
    print(f"无法自动打开地图: {str(e)}")
    print("请手动打开生成的HTML文件查看地图")

print("\n聚类分析完成！")
print("1. K-means聚类结果: kmeans_cluster_map.html")
print("2. DBSCAN聚类结果: dbscan_cluster_map.html")
print("3. 聚类统计图: kmeans_elbow_method.png")
print("4. 详细聚类数据: gps_clustering_results.xlsx") 
# GPS坐标聚类分析系统 (GPS Coordinate Clustering Analysis System)

## 项目概述 (Project Overview)

本项目是一个用于GPS坐标数据提取、可视化和聚类分析的系统。它可以从原始Excel文件中提取GPS坐标信息，使用多种可视化方法展示这些坐标点，并通过K-means和DBSCAN算法进行聚类分析，帮助识别地理位置的分布模式。

This project is a system for GPS coordinate extraction, visualization, and clustering analysis. It extracts GPS coordinate information from raw Excel files, displays these coordinates using various visualization methods, and performs clustering analysis using K-means and DBSCAN algorithms to help identify geographical distribution patterns.

## 主要功能 (Main Features)

1. **数据提取**：从Excel文件中提取GPS坐标、VIN码和事件时间信息
2. **数据可视化**：生成基础地图、热力图和聚类地图，直观展示GPS坐标分布
3. **聚类分析**：使用K-means和DBSCAN算法进行地理位置聚类
4. **结果导出**：将聚类结果保存为Excel文件和HTML交互式地图

1. **Data Extraction**: Extract GPS coordinates, VIN codes, and event time information from Excel files
2. **Data Visualization**: Generate base maps, heat maps, and cluster maps to visually display GPS coordinate distribution
3. **Clustering Analysis**: Perform geographical location clustering using K-means and DBSCAN algorithms
4. **Result Export**: Save clustering results as Excel files and HTML interactive maps

## 系统架构 (System Architecture)

本项目包含两个主要的Python脚本：

1. **extract_gps.py**: 数据提取脚本，从原始Excel文件中提取GPS坐标信息
2. **gps_clustering_analysis.py**: 聚类分析脚本，对提取的GPS坐标进行聚类分析
3. **visualize_gps_advanced.py**: 高级可视化脚本，生成多种类型的地图展示

This project contains two main Python scripts:

1. **extract_gps.py**: Data extraction script that extracts GPS coordinate information from raw Excel files
2. **gps_clustering_analysis.py**: Clustering analysis script that performs clustering analysis on extracted GPS coordinates
3. **visualize_gps_advanced.py**: Advanced visualization script that generates various types of maps

## 工作流程 (Workflow)

1. **数据提取阶段**：
   - 读取原始Excel文件（1.xlsx）
   - 从"事件地点"列中提取GPS坐标
   - 可选择性地提取"车辆VIN"和"事件时间"列
   - 将提取的数据保存为CSV和Excel格式

2. **数据可视化阶段**：
   - 读取提取的GPS坐标数据
   - 生成基础地图（base_map.html）
   - 生成热力图（heat_map.html）
   - 生成聚类地图（cluster_map.html）

3. **聚类分析阶段**：
   - 使用K-means算法进行聚类分析
   - 使用肘部法则和轮廓系数确定最佳聚类数量
   - 使用DBSCAN算法进行密度聚类
   - 生成聚类结果地图（kmeans_cluster_map.html和dbscan_cluster_map.html）
   - 将聚类结果保存为Excel文件（gps_clustering_results.xlsx）

1. **Data Extraction Phase**:
   - Read the original Excel file (1.xlsx)
   - Extract GPS coordinates from the "事件地点" (Event Location) column
   - Optionally extract "车辆VIN" (Vehicle VIN) and "事件时间" (Event Time) columns
   - Save extracted data in CSV and Excel formats

2. **Data Visualization Phase**:
   - Read the extracted GPS coordinate data
   - Generate a base map (base_map.html)
   - Generate a heat map (heat_map.html)
   - Generate a cluster map (cluster_map.html)

3. **Clustering Analysis Phase**:
   - Perform clustering analysis using the K-means algorithm
   - Determine the optimal number of clusters using the elbow method and silhouette coefficient
   - Perform density clustering using the DBSCAN algorithm
   - Generate clustering result maps (kmeans_cluster_map.html and dbscan_cluster_map.html)
   - Save clustering results as an Excel file (gps_clustering_results.xlsx)

## 输入输出文件 (Input/Output Files)

**输入文件**:
- `1.xlsx`: 原始Excel文件，包含事件地点、车辆VIN和事件时间信息

**中间文件**:
- `gps_coordinates.csv`: 提取的GPS坐标CSV文件
- `gps_coordinates.xlsx`: 提取的GPS坐标Excel文件

**输出文件**:
- `base_map.html`: 基础地图，显示所有GPS坐标点
- `heat_map.html`: 热力图，显示GPS坐标密度分布
- `cluster_map.html`: 聚类地图，显示GPS坐标的聚类情况
- `kmeans_cluster_map.html`: K-means聚类结果地图
- `dbscan_cluster_map.html`: DBSCAN聚类结果地图
- `kmeans_elbow_method.png`: K-means肘部法则和轮廓系数图
- `gps_clustering_results.xlsx`: 包含聚类标签的GPS坐标数据

**Input Files**:
- `1.xlsx`: Original Excel file containing event location, vehicle VIN, and event time information

**Intermediate Files**:
- `gps_coordinates.csv`: Extracted GPS coordinates in CSV format
- `gps_coordinates.xlsx`: Extracted GPS coordinates in Excel format

**Output Files**:
- `base_map.html`: Base map displaying all GPS coordinate points
- `heat_map.html`: Heat map showing GPS coordinate density distribution
- `cluster_map.html`: Cluster map showing GPS coordinate clustering
- `kmeans_cluster_map.html`: K-means clustering result map
- `dbscan_cluster_map.html`: DBSCAN clustering result map
- `kmeans_elbow_method.png`: K-means elbow method and silhouette coefficient graph
- `gps_clustering_results.xlsx`: GPS coordinate data with clustering labels

## 使用方法 (Usage)

1. 准备包含GPS坐标信息的Excel文件（命名为`1.xlsx`）
2. 运行数据提取脚本：`python extract_gps.py`
3. 运行聚类分析脚本：`visualize_gps_advanced.py` `python gps_clustering_analysis.py`
4. 查看生成的HTML地图文件和Excel结果文件

1. Prepare an Excel file containing GPS coordinate information (named `1.xlsx`)
2. Run the data extraction script: `python extract_gps.py`
3. Run the clustering analysis script: `visualize_gps_advanced.py` `python gps_clustering_analysis.py` 
4. View the generated HTML map files and Excel result files

## 技术栈 (Technology Stack)

- **Python**: 主要编程语言
- **Pandas**: 数据处理和分析
- **Scikit-learn**: 机器学习算法（K-means和DBSCAN）
- **Folium**: 交互式地图生成
- **Matplotlib**: 数据可视化

- **Python**: Main programming language
- **Pandas**: Data processing and analysis
- **Scikit-learn**: Machine learning algorithms (K-means and DBSCAN)
- **Folium**: Interactive map generation
- **Matplotlib**: Data visualization

## 环境要求 (Environment Requirements)

- Python 3.6+
- pandas
- numpy
- scikit-learn
- folium
- matplotlib
- openpyxl 
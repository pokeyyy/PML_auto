from typing import Any

import pandas as pd
import numpy as np
import math
from numpy import ndarray, dtype

def get_angle(p0, p1, p2):
    """
    计算圆弧的中心角，并以度为单位返回结果。

    参数：
        A, B, C: 三维空间中的点坐标，长度为3的可迭代对象。
            A, B 是圆弧的端点，C 是两条切线在端点处的交点。

    返回：
        圆心角（对应圆弧的角度，单位：度）。

    计算原理：
        1. 计算向量 CA, CB。
        2. 计算切线夹角 θ = ∠ACB = arccos((CA·CB)/(|CA||CB|))。
        3. 圆心角 φ = 180° - θ。
    """
    C = np.asarray(p0, dtype=float)
    A = np.asarray(p1, dtype=float)
    B = np.asarray(p2, dtype=float)

    # 向量 CA, CB
    v1 = A - C
    v2 = B - C

    # 计算切线夹角 θ
    cos_theta = np.dot(v1, v2) / (np.linalg.norm(v1) * np.linalg.norm(v2))
    cos_theta = np.clip(cos_theta, -1.0, 1.0)
    theta_rad = np.arccos(cos_theta)

    # 转换为度
    theta_deg = np.degrees(theta_rad)

    # 圆心角 φ（度）
    phi_deg = 180.0 - theta_deg
    return phi_deg


def find_P0(P1, P2, C):
    # 转换为 numpy 数组，方便计算
    P1 = np.array(P1, dtype=float)
    P2 = np.array(P2, dtype=float)
    C = np.array(C, dtype=float)

    # 计算向量 P1C 和 P2C
    P1C = P1 - C
    P2C = P2 - C

    # 计算法向量，平面法向量为 P1C 和 P2C 的外积
    normal = np.cross(P1C, P2C)

    # 如果法向量为零向量，说明 P1、P2、C 三点共线，无法计算平面法向量
    if np.allclose(normal, 0):
        raise ValueError("P1, P2, and C are collinear, no unique plane can be determined.")

    # 归一化法向量
    normal = normal / np.linalg.norm(normal)

    # 计算 P1 处切线的方向，切线方向垂直于 P1C 和法向量
    tangent_dir_P1 = np.cross(P1C, normal)
    tangent_dir_P1 = tangent_dir_P1 / np.linalg.norm(tangent_dir_P1)

    # 计算 P2 处切线的方向
    tangent_dir_P2 = np.cross(P2C, normal)
    tangent_dir_P2 = tangent_dir_P2 / np.linalg.norm(tangent_dir_P2)

    # 参数方程形式为 P1 + t * tangent_dir_P1 = P2 + s * tangent_dir_P2
    # 将它们设为相等求解 t 和 s，我们可以将它们联立成线性方程组

    # 构造矩阵 A 和向量 b
    A = np.vstack([tangent_dir_P1, -tangent_dir_P2]).T
    b = P2 - P1

    # 解线性方程组 A * [t, s] = b
    try:
        t_s = np.linalg.lstsq(A, b, rcond=None)[0]
    except np.linalg.LinAlgError:
        raise ValueError("Failed to solve the linear system, possibly due to numerical issues.")

    # 交点坐标
    intersection = P1 + t_s[0] * tangent_dir_P1

    return tuple(intersection)

def direction(p1, p2):
    p1 = np.array(p1)
    p2 = np.array(p2)
    p1p2 = p2 - p1
    dir1,dir2 = calc_angle(p1p2)
    return (dir1,dir2)

def orientation(type,p0,p1,p2 = None,p3 = None):
    if type=='elbow':
        p0 = np.array(p0)
        p1 = np.array(p1)
        p2 = np.array(p2)
        p1p0 = p0 - p1
        p2p0 = p0 - p2
        angle_x1,angle_x2 = calc_angle(p1p0)
        vector_z = np.cross(p1p0, p2p0)
        angle_z1,angle_z2 = calc_angle(vector_z)
        return (angle_x1,angle_x2,angle_z1,angle_z2)

    if type == 'tee':
        p0 = np.array(p0)
        p1 = np.array(p1)
        p3 = np.array(p3)
        p1p0 = p0 - p1
        angle_x1, angle_x2 = calc_angle(p1p0)
        p0p3 = p3 - p0
        angle_z1, angle_z2 = calc_angle(p0p3)
        return (angle_x1, angle_x2, angle_z1, angle_z2)

    if type == 'redu':
        p0 = np.array(p0)
        p1 = np.array(p1)
        p0p1 = p1 - p0
        angle_x1, angle_x2 = calc_angle(p0p1)
        return (angle_x1,angle_x2)


#计算向量相对于e,u轴的角度
def calc_angle(vector):
    projection_noe = vector[:-1]
    angle_rad = np.arctan2(projection_noe[1], projection_noe[0])
    angle = np.degrees(angle_rad)
    angle = (angle + 360) % 360
    # 计算向量的模
    magnitude = np.linalg.norm(vector)
    # 计算与 u 轴正方向的夹角
    angle_rad_u = np.arccos(vector[2] / magnitude)
    # 将弧度转换为度数
    angle_u = 90 - np.degrees(angle_rad_u)

    return angle,angle_u

#每个group结构为(type,id)
def read_group(location:str):
    df = pd.read_excel(location, sheet_name="Group")
    groups = []
    for row in range(df.shape[0]):
        group_str = df.iloc[row]["parts"]
        group = group_str.split("\n")
        group = [s[:-1] if i < len(group) - 1 else s for i, s in enumerate(group)]
        group = [eval(s) for s in group]
        groups.append(group)
    return groups


def move_points_along_line(p1, p2, distance):
    """
    移动p1和p2点到沿p1p2方向的指定距离。

    参数：
    p1: tuple (x1, y1, z1)，表示起始点的三维坐标。
    p2: tuple (x2, y2, z2)，表示终止点的三维坐标。
    distance: float，表示要移动的距离。

    返回值：
    tuple (new_p1, new_p2)，表示p1和p2沿p1p2方向移动后的新坐标。
    """
    # 计算方向向量
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    dz = p2[2] - p1[2]

    # 归一化方向向量
    length = math.sqrt(dx ** 2 + dy ** 2 + dz ** 2)
    if length == 0:
        raise ValueError("p1和p2不能是同一个点")
    unit_direction = (dx / length, dy / length, dz / length)

    # 计算p1的新坐标
    new_p1_x = p1[0] + unit_direction[0] * distance
    new_p1_y = p1[1] + unit_direction[1] * distance
    new_p1_z = p1[2] + unit_direction[2] * distance
    new_p1 = (new_p1_x, new_p1_y, new_p1_z)

    # 计算p2的新坐标
    new_p2_x = p2[0] + unit_direction[0] * distance
    new_p2_y = p2[1] + unit_direction[1] * distance
    new_p2_z = p2[2] + unit_direction[2] * distance
    new_p2 = (new_p2_x, new_p2_y, new_p2_z)

    return new_p1, new_p2

#对于[323.08437778372667, 379.3669199267068, 5.190823758635305]形式的转化
def transcoord(coord):
    coord = coord.split(",")
    coord[0] = float(coord[0].split('[')[1])
    coord[1] = float(coord[1])
    coord[2] = float(coord[2].split(']')[0])
    coord = np.asarray(coord)
    coord = coord * 1000
    return coord

def round_to_nearest(value):
    value = float(value)
    targets = [45, 90]
    closest = min(targets, key=lambda x: abs(x - value))
    return closest

def get_seq(path):
    sheets = pd.read_excel(path,sheet_name=None)

    # Gather component info
    components = []
    #anchors
    anchors = {}
    # cylinders
    for _, row in sheets['cylinders'].iterrows():
        components.append({
            'type': 'cylinder',
            'id': int(row['tid']),
            'anchors': {int(row['top_id']), int(row['bottom_id'])}
        })
    # elbows
    for _, row in sheets['elbows'].iterrows():
        components.append({
            'type': 'elbow',
            'id': int(row['tid']),
            'anchors': {int(row['p1_id']), int(row['p2_id'])}
        })
    # tees
    for _, row in sheets['tees'].iterrows():
        components.append({
            'type': 'tee',
            'id': int(row['tid']),
            'anchors': {int(row['top_1_id']), int(row['bottom_1_id'])}
        })
    # valves
    for _, row in sheets['valves'].iterrows():
        components.append({
            'type': 'valve',
            'id': int(row['tid']),
            'anchors': {int(row['p1_id']), int(row['p2_id'])}
        })
    # reducers
    for _, row in sheets['reducers'].iterrows():
        components.append({
            'type': 'reducer',
            'id': int(row['tid']),
            'anchors': {int(row['p1_id']), int(row['p2_id'])}
        })
    # flanges
    for _, row in sheets['flanges'].iterrows():
        components.append({
            'type': 'flange',
            'id': int(row['tid']),
            'anchors': {int(row['p1_id']), int(row['p2_id'])}
        })
    # gaskets
    for _, row in sheets['gaskets'].iterrows():
        components.append({
            'type': 'gasket',
            'id': int(row['tid']),
            'anchors': {int(row['p1_id']), int(row['p2_id'])}
        })
    #anchors
    for _, row in sheets['anchors'].iterrows():
        # 跳过含空值的行
        if pd.isna(row['comp1']) or pd.isna(row['tid_comp1']) or \
                pd.isna(row['comp2']) or pd.isna(row['tid_comp2']) or \
                pd.isna(row['tid']):
            continue  # 跳过这行

        comp1 = (row['comp1'], int(row['tid_comp1']))
        comp2 = (row['comp2'], int(row['tid_comp2']))
        id = int(row['tid'])
        anchors[id] = {comp1, comp2}

    #anchors索引
    comps_map = {
        (comp['type'], comp['id']): comp['anchors']
        for comp in components
    }

    # Traverse groups
    results_comp = []
    results_anchor = []
    for _, grp in sheets['groups'].iterrows():
        current_group = []
        group_anchor = []
        current_anchor = int(grp['top_anchors'])
        current_type = grp['top_comp']
        current_id = int(grp['top_id'])
        current_comp = (current_type, current_id)

        bottom_id = int(grp['bottom_id'])
        bottom_anchor = int(grp['bottom_anchors'])
        # Traverse path
        while True:
            current_group.append(current_comp)
            group_anchor.append(current_anchor)
            if(current_id == bottom_id):
                group_anchor.append(bottom_anchor)
                break
            current_anchors = comps_map[current_comp]
            other_anchor = (current_anchors - {current_anchor}).pop()
            other_comps = anchors[other_anchor]
            other_comp = (other_comps - {current_comp}).pop()

            current_anchor = other_anchor
            current_comp = other_comp
            current_id = current_comp[1]

        results_anchor.append(group_anchor)
        results_comp.append(current_group)
    return results_comp, results_anchor



# 测试
if __name__ == '__main__':
    path = 'resourse_new/E3D_2_Demo2.xlsx'
    res,res2 = get_seq(path)
    print(res)
    print(res2)
    print(len(res[0]))
    print(len(res2[0]))





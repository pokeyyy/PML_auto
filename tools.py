from typing import Any

import pandas as pd
import numpy as np
import math
from numpy import ndarray, dtype


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


# 创建branch
def create_branch(hpos,tpos,hdir,tdir,hbor,tbor,hstu):
    return {
        # 必填项：HPOS TPOS HDIR TDIR LHEA LTAI HBOR TBOR HCON DETA TEMP HSTU PSPE TCON
        "BUIL": "false",
        "SHOP": "false",
        "HPOS": f"E {hpos[0]:.3f}mm N {hpos[1]:.3f}mm U {hpos[2]:.3f}mm",
        "TPOS": f"E {tpos[0]:.3f}mm N {tpos[1]:.3f}mm U {tpos[2]:.3f}mm",
        "HDIR": f"E{hdir[0]:.3f}N{hdir[1]:.3f}U",
        "TDIR": f"E{tdir[0]:.3f}N{tdir[1]:.3f}U",
        "LHEA": "true",
        "LTAI": "true",
        "HBOR": f"{hbor}mm",
        "TBOR": f"{tbor}mm",
        "HCON": "OPEN",
        "TCON": "BWD",
        "LNTP": "unset",
        "TEMP": "-100000degC",  # 管道温度跟保温层厚度相关
        "PRES": "0pascal",
        "TPRESS": "0pascal",
        "HSTU": f"SPCOMPONENT {hstu}",
        "CCEN": "0",  # 默认为"0"
        "CCLA": "0",  # 默认为"0"
        "PSPE": "SPECIFICATION /NJU-SPEC",
        "DUTY": "'unset'",
        "DSCO": "'unset'",
        "PTSP": "'unset'",
        "INSC": "'unset'",
        "PTNB": "0",
        "PLANU": "unset",
        "DELDSG": "FALSE"
                  "\n"
    }

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

# 测试
if __name__ == '__main__':
    P1 = [40342.38556,13348.83315,3900]  # 第一个端点
    P2 = [40545.38556,13145.83315,3900]  # 第二个端点
    C = [40342.38556,13145.83315,3900]  # 圆心

    intersection = find_P0(P1, P2, C)
    print("两条切线的交点坐标为：", intersection)
    ori = orientation('elbow',intersection,P1,P2)
    print(ori)






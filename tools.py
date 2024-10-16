from typing import Any

import pandas as pd
import numpy as np
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

def orientation(type,p0,p1,p2,p3 = None):
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

# 测试
if __name__ == '__main__':
    P1 = [40342.38556,13348.83315,3900]  # 第一个端点
    P2 = [40545.38556,13145.83315,3900]  # 第二个端点
    C = [40342.38556,13145.83315,3900]  # 圆心

    intersection = find_P0(P1, P2, C)
    print("两条切线的交点坐标为：", intersection)
    ori = orientation('elbow',intersection,P1,P2)
    print(ori)






import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tools import orientation, find_P0, direction, move_points_along_line, transcoord, round_to_nearest, get_angle
from E3DModel import E3DModel, create_branch, create_PIPE, create_component_1, create_component_2, create_component_3, create_ZONE


# 创建主窗口
root = tk.Tk()
root.title("PML命令自生成程序")

# 设置窗口大小
root.geometry("400x280")

# 创建标签和输入框
tk.Label(root, text="输入框:").pack(pady=5)
base_file_path_entry = tk.Entry(root, width=50)
base_file_path_entry.pack(pady=5)

# 创建标签和输入框
tk.Label(root, text="输出框:").pack(pady=5)
folder_path_entry = tk.Entry(root, width=50)
folder_path_entry.pack(pady=5)


# 文件路径选择器
def select_base_file_path():
    paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    base_file_path_entry.delete(0, tk.END)
    base_file_path_entry.insert(0, ';'.join(paths))


def select_folder_path():
    path = filedialog.askdirectory()
    folder_path_entry.delete(0, tk.END)
    folder_path_entry.insert(0, path)


tk.Button(root, text="输入路径", command=select_base_file_path).pack(pady=5)
tk.Button(root, text="选择保存Excel文件路径", command=select_folder_path).pack(pady=5)

# 得到anchor的坐标
def findposition(id):
    base_file_paths = base_file_path_entry.get().split(';')
    df = pd.read_excel(base_file_paths[0], sheet_name="anchors")
    anchor = df.loc[df["tid"] == id]
    return (anchor.iloc[0]["coord1"],anchor.iloc[0]["coord2"],anchor.iloc[0]["coord3"])

def process_data():
    base_file_paths = base_file_path_entry.get().split(';')
    folder = folder_path_entry.get()

    if not base_file_paths or not folder:
        messagebox.showerror("输入错误", "请填写所有路径")
        return

    # 输出目录
    output_dir = folder
    os.makedirs(output_dir, exist_ok=True)

    # 创建根层级SITE
    # 使用一个字典存储当前层级的计数
    level_counts = {
        'SITE': 1,
        'ZONE': 1,
        'PIPE': 1,
        'BRANCH': 1,
        'COMPONENT': 1
    }

    # 根层级SITE
    site = E3DModel("SITE", f"/GL1001\n")
    level_counts['SITE'] += 1

    # 层级存储 字典
    hierarchy = {'SITE': site}

    # 创建ZONE层级
    zone_name = f"/ZONE{level_counts['ZONE']}\n"
    zone_attributes = create_ZONE("PD")
    zone = E3DModel("ZONE", zone_name, zone_attributes)
    level_counts['ZONE'] += 1
    hierarchy['SITE'].add_child(zone)
    hierarchy['ZONE'] = zone

    # 创建PIPE层级
    pipe_name = "/200-GH-R2271-RL1A-N"
    pip_attributes = create_PIPE("PD")
    pipe = E3DModel("PIPE", pipe_name, pip_attributes)
    level_counts['PIPE'] += 1
    if 'ZONE' in hierarchy:
        hierarchy['ZONE'].add_child(pipe)
    hierarchy['PIPE'] = pipe


    for base_file_path in base_file_paths:
        sheets = pd.read_excel(base_file_path,sheet_name=None)
        for sheet_name, df in sheets.items():
            if sheet_name == "anchors": continue
            elif sheet_name == "elbows":
                for index, row in df.iterrows():
                    # 为每个component创建一个branch
                    branch_name = f"{pipe_name}/B{level_counts['BRANCH']}"

                    row_dict = row.to_dict()

                    hpos = findposition(row_dict["p1_id"])
                    tpos = findposition(row_dict["p2_id"])
                    center = findposition(row_dict["center_id"])
                    POS = find_P0(hpos, tpos, center)
                    hdir = direction(hpos,POS)
                    tdir = direction(tpos,POS)
                    hbor = row_dict["Processed Diameter_DN"]
                    tbor = hbor
                    hstu = map_name("tube", hbor)
                    branch_attributes = create_branch(hpos,tpos, hdir, tdir,hbor,tbor,hstu,"PD")
                    branch = E3DModel("BRANCH", branch_name,branch_attributes)
                    level_counts['BRANCH'] += 1
                    if 'PIPE' in hierarchy:
                        hierarchy['PIPE'].add_child(branch)
                    hierarchy['BRANCH'] = branch

                    p1 = hpos
                    p2 = tpos
                    ORI = orientation("elbow", POS, p1, p2)
                    bore = row_dict["Processed Diameter_DN"]
                    angle = get_angle(POS, p1, p2)
                    angle = round_to_nearest(angle)
                    SPRE = map_name("elbow", bore, angle=angle)
                    if SPRE is None:
                        continue
                    LSTU = map_name("tube", bore)

                    elbow_attributes = create_component_1(POS,ORI,SPRE,LSTU)
                    component_name = f"/ELBOW-{level_counts['COMPONENT']}"
                    component = E3DModel("ELBOW", component_name, elbow_attributes)
                    level_counts['COMPONENT'] += 1
                    if 'BRANCH' in hierarchy:
                        hierarchy['BRANCH'].add_child(component)
                    hierarchy['COMPONENT'] = component

            elif sheet_name == "tees":
                for index, row in df.iterrows():
                    branch_name = f"{pipe_name}/B{level_counts['BRANCH']}"

                    row_dict = row.to_dict()
                    hpos = findposition(row_dict["top_1_id"])
                    tpos = findposition(row_dict["bottom_1_id"])
                    hdir = direction(hpos,tpos)
                    tdir = direction(tpos,hpos)
                    hbor = row_dict["Processed Diameter_DN1"]
                    tbor = hbor
                    hstu = map_name("tube", hbor)
                    branch_attributes = create_branch(hpos, tpos, hdir, tdir, hbor, tbor, hstu,"PD")
                    branch = E3DModel("BRANCH", branch_name, branch_attributes)
                    level_counts['BRANCH'] += 1
                    if 'PIPE' in hierarchy:
                        hierarchy['PIPE'].add_child(branch)
                    hierarchy['BRANCH'] = branch

                    POS = findposition(row_dict["bottom_2_id"])
                    p1 = hpos
                    p2 = tpos
                    p3 = findposition(row_dict["top_2_id"])
                    ORI = orientation('tee', POS, p1, p2, p3)
                    bore1 = row_dict["Processed Diameter_DN1"]
                    bore2 = row_dict["Processed Diameter_DN2"]
                    SPRE = map_name("tee", bore1, bore2)
                    LSTU = map_name("tube", bore1)

                    tee_attributes = create_component_1(POS,ORI,SPRE,LSTU)
                    component_name = f"/TEE-{level_counts['COMPONENT']}"
                    component = E3DModel("TEE", component_name, tee_attributes)
                    level_counts['COMPONENT'] += 1
                    if 'BRANCH' in hierarchy:
                        hierarchy['BRANCH'].add_child(component)
                    hierarchy['COMPONENT'] = component

            elif sheet_name == "cylinders":
                for index, row in df.iterrows():
                    branch_name = f"{pipe_name}/B{level_counts['BRANCH']}"

                    row_dict = row.to_dict()
                    hpos = findposition(row_dict["top_id"])
                    tpos = findposition(row_dict["bottom_id"])
                    hdir = direction(hpos,tpos)
                    tdir = direction(tpos,hpos)
                    hbor = row_dict["Processed Diameter_DN"]
                    tbor = hbor
                    hstu = map_name("tube", hbor)
                    branch_attributes = create_branch(hpos, tpos, hdir, tdir, hbor, tbor, hstu,"PD")
                    branch = E3DModel("BRANCH", branch_name, branch_attributes)
                    level_counts['BRANCH'] += 1
                    if 'PIPE' in hierarchy:
                        hierarchy['PIPE'].add_child(branch)
                    hierarchy['BRANCH'] = branch

            elif sheet_name == "reducers":
                for index, row in df.iterrows():
                    branch_name = f"{pipe_name}/B{level_counts['BRANCH']}"

                    row_dict = row.to_dict()
                    hpos = findposition(row_dict["p1_id"])
                    tpos = findposition(row_dict["p2_id"])
                    hdir = direction(hpos, tpos)
                    tdir = direction(tpos, hpos)
                    hbor = int(row_dict["Processed Diameter_DN1"])
                    tbor = int(row_dict["Processed Diameter_DN2"])
                    hstu = map_name("tube", hbor)
                    branch_attributes = create_branch(hpos, tpos, hdir, tdir, hbor, tbor, hstu,"PD")
                    branch = E3DModel("BRANCH", branch_name, branch_attributes)
                    level_counts['BRANCH'] += 1
                    if 'PIPE' in hierarchy:
                        hierarchy['PIPE'].add_child(branch)
                    hierarchy['BRANCH'] = branch

                    POS = hpos
                    p2 = tpos
                    bore1 = hbor
                    bore2 = tbor
                    ORI = orientation('redu', POS, p2)
                    SPRE = map_name("redu", bore1, bore2)
                    LSTU = map_name("tube", bore2)

                    redu_attributes = create_component_2(POS,ORI,SPRE,LSTU)
                    component_name = f"/REDU-{level_counts['COMPONENT']}"
                    component = E3DModel("REDUCER", component_name, redu_attributes)
                    level_counts['COMPONENT'] += 1
                    if 'BRANCH' in hierarchy:
                        hierarchy['BRANCH'].add_child(component)
                    hierarchy['COMPONENT'] = component

            elif sheet_name == "valves":
                for index, row in df.iterrows():
                    branch_name = f"{pipe_name}/B{level_counts['BRANCH']}"

                    row_dict = row.to_dict()
                    hpos = findposition(row_dict["p1_id"])
                    tpos = findposition(row_dict["p2_id"])
                    hdir = direction(hpos, tpos)
                    tdir = direction(tpos, hpos)
                    bore = row_dict["Processed Diameter_DN"]
                    hstu = map_name("tube", bore)
                    branch_attributes = create_branch(hpos, tpos, hdir, tdir, bore, bore, hstu,"PD")
                    branch = E3DModel("BRANCH", branch_name, branch_attributes)
                    level_counts['BRANCH'] += 1
                    if 'PIPE' in hierarchy:
                        hierarchy['PIPE'].add_child(branch)
                    hierarchy['BRANCH'] = branch

                    POS = tuple((x + y) / 2 for x, y in zip(hpos, tpos))
                    ORI = orientation('redu', POS, tpos)
                    SPRE = map_name("valv", bore)
                    LSTU = hstu

                    valv_attributes = create_component_2(POS,ORI,SPRE,LSTU)

                    component_name = f"/VALV-{level_counts['COMPONENT']}"
                    component = E3DModel("VALVE", component_name, valv_attributes)
                    level_counts['COMPONENT'] += 1
                    if 'BRANCH' in hierarchy:
                        hierarchy['BRANCH'].add_child(component)
                    hierarchy['COMPONENT'] = component

            elif sheet_name == "flanges":
                for i in range(0, len(df), 2):
                    rows = df.iloc[i:i + 2]
                    row1 = rows.iloc[0]
                    row2 = rows.iloc[1]
                    row1_dict = row1.to_dict()
                    row2_dict = row2.to_dict()
                    branch_name = f"{pipe_name}/B{level_counts['BRANCH']}"

                    r2_0 = findposition(row2_dict["p1_id"])
                    r2_2 = findposition(row2_dict["p2_id"])
                    r2_0,r2_2 = move_points_along_line(r2_0, r2_2, 4.5)

                    hpos = findposition(row1_dict["p2_id"])
                    tpos = r2_2
                    hdir = direction(hpos, tpos)
                    tdir = direction(tpos, hpos)
                    hbor = int(row1_dict["Processed Diameter_DN"])
                    tbor = int(row2_dict["Processed Diameter_DN"])
                    hstu = map_name("tube", hbor)
                    branch_attributes = create_branch(hpos, tpos, hdir, tdir, hbor, tbor, hstu,"PD")
                    branch = E3DModel("BRANCH", branch_name, branch_attributes)
                    level_counts['BRANCH'] += 1
                    if 'PIPE' in hierarchy:
                        hierarchy['PIPE'].add_child(branch)
                    hierarchy['BRANCH'] = branch

                    POS = findposition(row1_dict["p1_id"])
                    p2 = hpos
                    ORI = orientation('redu', POS, p2)
                    bore = hbor
                    SPRE = map_name("flan", bore)
                    LSTU = map_name("tube", bore)
                    flan_attributes = create_component_3(POS,ORI,SPRE,LSTU,2,1)

                    component_name = f"/FLAN-{level_counts['COMPONENT']}"
                    component = E3DModel("FLANGE", component_name, flan_attributes)
                    level_counts['COMPONENT'] += 1
                    if 'BRANCH' in hierarchy:
                        hierarchy['BRANCH'].add_child(component)
                    hierarchy['COMPONENT'] = component


                    p2 = r2_0
                    ORI = orientation('redu',POS, p2)
                    SPRE = map_name("gask", bore)
                    gasket_attributes = create_component_3(POS,ORI,SPRE,LSTU,2,1)

                    component_name = f"/GASK-{level_counts['COMPONENT']}"
                    component = E3DModel("GASKET", component_name, gasket_attributes)
                    level_counts['COMPONENT'] += 1
                    if 'BRANCH' in hierarchy:
                        hierarchy['BRANCH'].add_child(component)
                    hierarchy['COMPONENT'] = component

                    POS = r2_0
                    p2 = r2_2
                    ORI = orientation('redu', POS, p2)
                    bore = tbor
                    SPRE = map_name("flan", bore)
                    LSTU = map_name("tube", bore)
                    flan_attributes = create_component_3(POS,ORI,SPRE,LSTU,1,2)

                    component_name = f"/FLAN-{level_counts['COMPONENT']}"
                    component = E3DModel("FLANGE", component_name, flan_attributes)
                    level_counts['COMPONENT'] += 1
                    if 'BRANCH' in hierarchy:
                        hierarchy['BRANCH'].add_child(component)
                    hierarchy['COMPONENT'] = component

    # 生成命令
    commands = site.generate_commands()

    # 保存修改后的txt文件
    output_file = os.path.join(output_dir, "e3d_commands.txt")
    with open(output_file, "w") as file:
        for command in commands:
            file.write(command + "\n")

    # 输出命令
    for command in commands:
        print(command)


# 映射dn值对应的名称，三通,大小头需要两个bore
# 初始化失败计数器
fail_count = {
    'tube': 0,
    'elbow': 0,
    'tee': 0,
    'redu': 0,
    'flan': 0,
    'gask': 0,
    'valv': 0
}

def map_name(type, bore1, bore2=None, angle=None):
    try:
        if type == 'tube':
            return dic_tube[bore1]
        elif type == 'elbow':
            return dic_elbow[(bore1, angle)]
        elif type == 'tee':
            return dic_tee[(bore1, bore2)]
        elif type == 'redu':
            return dic_redu[(bore1, bore2)]
        elif type == 'flan':
            return dic_flan[bore1]
        elif type == 'gask':
            return dic_gask[bore1]
        elif type == 'valv':
            return dic_valv[bore1]
    except KeyError:
        fail_count[type] += 1
        print(f"[FAIL] Type: {type}, Bore1: {bore1}, Bore2: {bore2}, Angle: {angle}")
        return None


 # 预处理：创建dn对应字典
sheets_ref = pd.read_excel(".\\resourse_new\\reflaction.xlsx", sheet_name=None, header=None)
df_elbow = sheets_ref['elbow']
df_tube = sheets_ref["tube"]
df_tee = sheets_ref["tee"]
df_redu = sheets_ref["redu"]
df_flan = sheets_ref["flan"]
df_gask = sheets_ref["gask"]
df_valv = sheets_ref["valv"]
dic_tube = dict(zip(df_tube.iloc[:, 2], df_tube.iloc[:, 0]))
dic_elbow = dict(zip(zip(df_elbow.iloc[:, 2], df_elbow.iloc[:, 3]), df_elbow.iloc[:, 0]))
dic_tee = dict(zip(zip(df_tee.iloc[:, 2], df_tee.iloc[:, 3]), df_tee.iloc[:, 0]))
dic_redu = dict(zip(zip(df_redu.iloc[:, 2], df_redu.iloc[:, 3]), df_redu.iloc[:, 0]))
dic_flan = dict(zip(df_flan.iloc[:, 2], df_flan.iloc[:, 0]))
dic_gask = dict(zip(df_gask.iloc[:, 2], df_gask.iloc[:, 0]))
dic_valv = dict(zip(df_valv.iloc[:, 2], df_valv.iloc[:, 0]))
print(len(dic_elbow))

# 创建处理数据按钮
tk.Button(root, text="点一下|自动出来PML命令！", command=process_data).pack(pady=20)

# 运行主窗口
root.mainloop()

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tools import orientation, find_P0, direction, move_points_along_line, transcoord, round_to_nearest
from E3DModel import E3DModel, create_branch, create_PIPE, create_component_1, create_component_2, create_component_3

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

#记录branch
branchs = {}

# 得到anchor的坐标
def findposition(id):
    base_file_paths = base_file_path_entry.get().split(';')
    df = pd.read_excel(base_file_paths[0], sheet_name="Anchor Parameters")
    anchor = df.loc[df["tid"] == id]
    print(anchor.iloc[0]["coord1"])
    print(anchor.iloc[0]["coord2"])
    print(anchor.iloc[0]["coord3"])
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
    site = E3DModel("SITE", f"/Dikaer-TEST-{level_counts['SITE']}\n")
    level_counts['SITE'] += 1

    # 层级存储 字典
    hierarchy = {'SITE': site}

    # 创建ZONE层级
    zone_name = f"/ZONE{level_counts['ZONE']}\n"
    zone = E3DModel("ZONE", zone_name)
    level_counts['ZONE'] += 1
    hierarchy['SITE'].add_child(zone)
    hierarchy['ZONE'] = zone

    # 创建PIPE层级
    pipe_name = f"/PIPE{level_counts['PIPE']}"
    pip_attributes = create_PIPE()
    pipe = E3DModel("PIPE", pipe_name, pip_attributes)
    level_counts['PIPE'] += 1
    if 'ZONE' in hierarchy:
        hierarchy['ZONE'].add_child(pipe)
    hierarchy['PIPE'] = pipe


    for base_file_path in base_file_paths:
        sheets = pd.read_excel(base_file_path,sheet_name=None)
        for sheet_name, df in sheets.items():
            if sheet_name == "elbows":
                for index, row in df.iterrows():
                    # 创建component
                    row_dict = row.to_dict()
                    if(row_dict["Processed Diameter_DN"] == -1):
                        continue

                    p1 = transcoord(row_dict["coord1"])
                    p2 = transcoord(row_dict["coord2"])
                    center = transcoord(row_dict["center_coord"])
                    p0 = find_P0(p1, p2, center)
                    ORI = orientation("elbow", p0, p1, p2)
                    bore = row_dict["Processed Diameter_DN"]
                    angle = round_to_nearest(row_dict["angle"])
                    branch_id = int(row_dict["group_id"])

                    SPRE = map_name("elbow", bore, angle=angle) if map_name("elbow", bore, angle=angle) != -1 else -1
                    LSTU = map_name("tube", bore) if map_name("tube", bore, angle=angle) != -1 else -1
                    if SPRE == -1 or LSTU == -1:
                        continue

                    elbow_attributes = create_component_1(p0,ORI,SPRE,LSTU)
                    component_name = f"/ELBOW-{level_counts['COMPONENT']}"
                    component = E3DModel("ELBOW", component_name, elbow_attributes)
                    level_counts['COMPONENT'] += 1
                    #判断是否需要创建branch
                    if branch_id == -1:
                        branch_name = f"/PIPE1/fail/B{level_counts['BRANCH']}"
                        hpos = p1
                        tpos = p2
                        POS = p0
                        hdir = direction(hpos, POS)
                        tdir = direction(tpos, POS)
                        hbor = row_dict["Processed Diameter_DN"]
                        tbor = hbor
                        hstu = map_name("tube", hbor)
                        branch_attributes = create_branch(hpos, tpos, hdir, tdir, hbor, tbor, hstu)
                        branch = E3DModel("BRANCH", branch_name, branch_attributes)
                        level_counts['BRANCH'] += 1
                        if 'PIPE' in hierarchy:
                            hierarchy['PIPE'].add_child(branch)
                        hierarchy['BRANCH'] = branch
                        hierarchy['BRANCH'].add_child(component)
                        hierarchy['COMPONENT'] = component
                        branchs[branch_id] = branch

                    else:
                        if branch_id in branchs :
                            branchs[branch_id].add_child(component)
                        else:
                            branch_name = f"/PIPE1/B{branch_id}"
                            hpos = p1
                            tpos = p2
                            POS = p0
                            hdir = direction(hpos, POS)
                            tdir = direction(tpos, POS)
                            hbor = row_dict["Processed Diameter_DN"]
                            tbor = hbor
                            hstu = map_name("tube", hbor)
                            branch_attributes = create_branch(hpos, tpos, hdir, tdir, hbor, tbor, hstu)
                            branch = E3DModel("BRANCH", branch_name, branch_attributes)

                            level_counts['BRANCH'] += 1
                            if 'PIPE' in hierarchy:
                                hierarchy['PIPE'].add_child(branch)
                            hierarchy['BRANCH'] = branch
                            hierarchy['BRANCH'].add_child(component)
                            hierarchy['COMPONENT'] = component
                            branchs[branch_id] = branch

            elif sheet_name == "tees":
                for index, row in df.iterrows():
                    row_dict = row.to_dict()
                    if row_dict["Processed Diameter_DN1"] == -1 or row_dict["Processed Diameter_DN2"] == -1 :
                        continue

                    POS = transcoord(row_dict["bottom2"])
                    p1 = transcoord(row_dict["top1"])
                    p2 = transcoord(row_dict["bottom1"])
                    p3 = transcoord(row_dict["top2"])
                    ORI = orientation('tee', POS, p1, p2, p3)
                    bore1 = row_dict["Processed Diameter_DN1"]
                    bore2 = row_dict["Processed Diameter_DN2"]
                    SPRE = map_name("tee", bore1, bore2) if map_name("tee", bore1, bore2) != -1 else -1
                    LSTU = map_name("tube", bore1) if map_name("tube", bore1) != -1 else -1
                    if SPRE == -1 or LSTU == -1:
                        continue
                    branch_id = int(row_dict["group_id"])

                    tee_attributes = create_component_1(POS,ORI,SPRE,LSTU)
                    component_name = f"/TEE-{level_counts['COMPONENT']}"
                    component = E3DModel("TEE", component_name, tee_attributes)
                    level_counts['COMPONENT'] += 1

                    if branch_id == -1:
                        branch_name = f"/PIPE1/fail/B{level_counts['BRANCH']}"
                        hpos = p1
                        tpos = p2
                        hdir = direction(hpos, tpos)
                        tdir = direction(tpos, hpos)
                        hbor = row_dict["Processed Diameter_DN1"]
                        tbor = hbor
                        hstu = map_name("tube", hbor)
                        branch_attributes = create_branch(hpos, tpos, hdir, tdir, hbor, tbor, hstu)
                        branch = E3DModel("BRANCH", branch_name, branch_attributes)
                        level_counts['BRANCH'] += 1
                        if 'PIPE' in hierarchy:
                            hierarchy['PIPE'].add_child(branch)
                        hierarchy['BRANCH'] = branch
                        hierarchy['BRANCH'].add_child(component)
                        hierarchy['COMPONENT'] = component
                        branchs[branch_id] = branch

                    else:
                        if branch_id in branchs:
                            branchs[branch_id].add_child(component)
                        else:
                            branch_name = f"/PIPE1/B{branch_id}"
                            hpos = p1
                            tpos = p2
                            hdir = direction(hpos, tpos)
                            tdir = direction(tpos, hpos)
                            hbor = row_dict["Processed Diameter_DN"]
                            tbor = hbor
                            hstu = map_name("tube", hbor)
                            branch_attributes = create_branch(hpos, tpos, hdir, tdir, hbor, tbor, hstu)
                            branch = E3DModel("BRANCH", branch_name, branch_attributes)
                            level_counts['BRANCH'] += 1
                            if 'PIPE' in hierarchy:
                                hierarchy['PIPE'].add_child(branch)
                            hierarchy['BRANCH'] = branch
                            hierarchy['BRANCH'].add_child(component)
                            hierarchy['COMPONENT'] = component
                            branchs[branch_id] = branch

            elif sheet_name == "cylinders":
                for index, row in df.iterrows():
                    row_dict = row.to_dict()
                    if (row_dict["Processed Diameter_DN"] == -1):
                        continue
                    branch_id = row_dict["group_id"]
                    if branch_id == -1:
                        branch_name = f"/PIPE1/fail/B{level_counts['BRANCH']}"
                        hpos = transcoord(row_dict["top_center"])
                        tpos = transcoord(row_dict["bottom_center"])
                        hdir = direction(hpos, tpos)
                        tdir = direction(tpos, hpos)
                        hbor = row_dict["Processed Diameter_DN"]
                        tbor = hbor
                        hstu = map_name("tube", hbor) if map_name("tube", hbor) != -1 else -1
                        if hstu == -1:
                            continue

                        branch_attributes = create_branch(hpos, tpos, hdir, tdir, hbor, tbor, hstu)
                        branch = E3DModel("BRANCH", branch_name, branch_attributes)
                        level_counts['BRANCH'] += 1
                        if 'PIPE' in hierarchy:
                            hierarchy['PIPE'].add_child(branch)
                        hierarchy['BRANCH'] = branch

                    else:
                        if branch_id in branchs:
                            continue
                        else:
                            branch_name = f"/PIPE1/B{branch_id}"
                            hpos = transcoord(row_dict["top_center"])
                            tpos = transcoord(row_dict["bottom_center"])
                            hdir = direction(hpos, tpos)
                            tdir = direction(tpos, hpos)
                            hbor = row_dict["Processed Diameter_DN"]
                            tbor = hbor
                            hstu = map_name("tube", hbor) if map_name("tube", hbor) != -1 else -1
                            if hstu == -1:
                                continue
                            branch_attributes = create_branch(hpos, tpos, hdir, tdir, hbor, tbor, hstu)
                            branch = E3DModel("BRANCH", branch_name, branch_attributes)
                            level_counts['BRANCH'] += 1
                            if 'PIPE' in hierarchy:
                                hierarchy['PIPE'].add_child(branch)
                            hierarchy['BRANCH'] = branch
                            branchs[branch_id] = branch

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
def map_name(type, bore1, bore2=None, angle=None):
    if type == 'tube':
        if bore1 in dic_tube:
            return dic_tube[bore1]
        else:
            print(f"{bore1} not in dic_tube")
            return -1
    elif type == 'elbow':
        if (bore1,angle) in dic_elbow:
            return dic_elbow[(bore1, angle)]
        else:
            print(f"{bore1,angle} not in dic_elbow")
            return -1
    elif type == 'tee':
        if (bore1,bore2) in dic_tee:
            return dic_tee[(bore1, bore2)]
        else:
            print(f"{bore1,bore2} not in dic_tee")
            return -1
    elif type == 'redu':
        return dic_redu[(bore1, bore2)]
    elif type == 'flan':
        return dic_flan[bore1]
    elif type == 'gask':
        return dic_gask[bore1]
    elif type == 'valv':
        return dic_valv[bore1]


 # 预处理：创建dn对应字典
sheets_ref = pd.read_excel(".\\resourse\\reflaction.xlsx", sheet_name=None)
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

# 创建处理数据按钮
tk.Button(root, text="点一下|自动出来PML命令！", command=process_data).pack(pady=20)

# 运行主窗口
root.mainloop()

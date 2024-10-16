import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tools import orientation, find_P0

"""
根据Excel内首行title类型[title修改行：113]判断component类型(南大打组案例Waiting)
注意：管件坐标值尚未填写（所有必填命令同理），生成的PML命令不完整
"""


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

    class E3DModel:
        def __init__(self, model_type, name, attributes=None, children=None):
            self.model_type = model_type
            self.name = name
            self.attributes = attributes if attributes else {}
            self.children = children if children else []

        def add_attribute(self, key, value):
            self.attributes[key] = value

        def add_child(self, child):
            self.children.append(child)

        # 左对齐开关=left_align，调试的时候用
        def generate_commands(self, indent=0, left_align = True):
            indent_str = '' if left_align else ' ' * indent
            commands = []

            commands.append(f"{indent_str}NEW {self.model_type} {self.name}")

            for key, value in self.attributes.items():
                commands.append(f"{indent_str}{key} {value}")

            for child in self.children:
                commands.extend(child.generate_commands(indent + 4, left_align))

            commands.append(f"{indent_str}END")
            return commands

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
    site = E3DModel("SITE", f"/Dikaer-TEST2-{level_counts['SITE']}\n")
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
    pipe = E3DModel("PIPE", pipe_name, {
        # 必填项：TEMP PSPE BORE
        "BUIL": "false",
        "SHOP": "false",
        "TEMP": "-100000degC",  # 管道温度跟保温层厚度相关
        "PRES": "0pascal",
        "TPRESS": "0pascal",
        "PSPE": "SPECIFICATION /NJU-SPEC",
        "CCEN": "0",
        "CCLA": "0",
        "LNTP": "unset",
        "BORE": "100mm",
        "DUTY": "'unset'",
        "DSCO": "'unset'",
        "PTSP": "'unset'",
        "INSC": "'unset'",
        "DELDSG": "FALSE",
        "PLANU": "unset"
                     "\n"
        })
    level_counts['PIPE'] += 1
    if 'ZONE' in hierarchy:
        hierarchy['ZONE'].add_child(pipe)
    hierarchy['PIPE'] = pipe


    for base_file_path in base_file_paths:
        sheets = pd.read_excel(base_file_path,sheet_name=None)
        df_elbow_para = sheets["Elbow Parameters"]

        branch_head = df_elbow_para.iloc[0]["p1_id"]
        branch_tail = df_elbow_para.iloc[-1]["p2_id"]
        hpos = findposition(branch_head)
        tpos = findposition(branch_tail)
        hbor = df_elbow_para.iloc[0]["Processed Diameter_DN"]
        tbor = df_elbow_para.iloc[-1]["Processed Diameter_DN"]
        hstu = map_name("tube",hbor)
        branch_name = f"/B{level_counts['BRANCH']}/B{level_counts['BRANCH']}"
        branch = E3DModel("BRANCH", branch_name, {
            # 必填项：HPOS TPOS HDIR TDIR LHEA LTAI HBOR TBOR HCON DETA TEMP HSTU PSPE TCON
            "BUIL": "false",
            "SHOP": "false",
            "HPOS": f"E {hpos[0]:.3f}mm N {hpos[1]:.3f}mm U {hpos[2]:.3f}mm",
            "TPOS": f"E {tpos[0]:.3f}mm S {tpos[1]:.3f}mm U {tpos[2]:.3f}mm",
            "HDIR": "N",
            "TDIR": "N",
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
        })
        level_counts['BRANCH'] += 1
        if 'PIPE' in hierarchy:
            hierarchy['PIPE'].add_child(branch)
        hierarchy['BRANCH'] = branch

        for sheet_name, df in sheets.items():
            if sheet_name == "Anchor Parameters": continue

            type=""
            # 创建COMPONENT层级
            if sheet_name == "Elbow Parameters":
                print(df)
                type = "ELBOW"
                for index, row in df.iterrows():
                    row_dict = row.to_dict()
                    center = findposition(row_dict["center_id"])
                    p1 = findposition(row_dict["p1_id"])
                    p2 = findposition(row_dict["p2_id"])
                    POS = find_P0(p1,p2,center)
                    ORI = orientation("elbow",POS, p1, p2)
                    bore = row_dict["Processed Diameter_DN"]
                    angle = row_dict["angle"]
                    SPRE = map_name("elbow",bore,angle = angle)
                    LSTU = map_name("tube",bore)

                    elbow_attributes = {
                        "POS": f"E {POS[0]:.3f}mm N {POS[1]:.3f}mm U {POS[2]:.3f}mm",  # 必填数据——弯头元件的坐标值
                        "ORI": f"X is E{round(ORI[0])}N{round(ORI[1])}U and Z is E{round(ORI[2])}N{round(ORI[3])}U",  # 必填数据——弯头元件的朝向
                        "BUIl": "false",
                        "SHOP": "true",
                        "SPRE": f"SPCOMPONENT {SPRE}",  # 必填数据——管道等级中对应弯头元件的name
                        "LSTU": f"SPCOMPONENT {LSTU}",  # 必填数据——管道等级中对应管道的元件name
                        "ORIF": "true",
                        "POSF": "true"
                                "\n"
                    }

                    component_name = f"/ELBOW-{level_counts['COMPONENT']}"
                    component = E3DModel("ELBOW", component_name, elbow_attributes)
                    level_counts['COMPONENT'] += 1
                    if 'BRANCH' in hierarchy:
                        hierarchy['BRANCH'].add_child(component)
                    hierarchy['COMPONENT'] = component

            elif sheet_name == "Tee Parameters":
                type = "TEE"
                for index, row in df.iterrows():
                    row_dict = row.to_dict()
                    POS = findposition(row_dict["bottom_2_id"])
                    p1 = findposition(row_dict["top_1_id"])
                    p2 = findposition(row_dict["bottom_1_id"])
                    p3 = findposition(row_dict["top_2_id"])
                    ORI = orientation('tee',POS, p1,p2,p3)
                    bore1 = row_dict["Processed Diameter_DN1"]
                    bore2 = row_dict["Processed Diameter_DN2"]
                    SPRE = map_name("tee", bore1, bore2)
                    LSTU = map_name("tube", bore1)

                    tee_attributes = {
                        "POS": f"E {POS[0]}mm N {POS[1]}mm U {POS[2]}mm",  # 必填数据——三通元件的坐标值
                        "ORI": f"X is E{round(ORI[0])}N{round(ORI[1])}U and Z is E{round(ORI[2])}N{round(ORI[3])}U",  # 必填数据——弯头元件的朝向
                        "BUIl": "false",
                        "SHOP": "true",
                        "SPRE": f"SPCOMPONENT {SPRE}",  # 必填数据——管道等级中对应弯头元件的name
                        "LSTU": f"SPCOMPONENT {LSTU}",  # 必填数据——管道等级中对应管道的元件name
                        "ORIF": "true",
                        "POSF": "true"
                                "\n"
                    }
                    component_name = f"/TEE-{level_counts['COMPONENT']}"
                    component = E3DModel("TEE", component_name, tee_attributes)
                    level_counts['COMPONENT'] += 1
                    if 'BRANCH' in hierarchy:
                        hierarchy['BRANCH'].add_child(component)
                    hierarchy['COMPONENT'] = component


    # 生成命令
    commands = site.generate_commands()

    # 保存修改后的txt文件
    output_file = os.path.join(output_dir, "e3d_commands_branch.txt")
    with open(output_file, "w") as file:
        for command in commands:
            file.write(command + "\n")

    # 输出命令
    for command in commands:
        print(command)


# 映射dn值对应的名称，不需要angle的元件angle为-1，三通需要两个bore
def map_name(type, bore1, bore2=None, angle=None):
    if type == 'tube':
        return dic_tube[bore1]
    elif type == 'elbow':
        return dic_elbow[(bore1, angle)]
    elif type == 'tee':
        return dic_tee[(bore1, bore2)]

 # 预处理：创建dn对应字典
sheets_ref = pd.read_excel(".\\resourse\\reflaction.xlsx", sheet_name=None)
df_elbow = sheets_ref['elbow']
df_tube = sheets_ref["tube"]
df_tee = sheets_ref["tee"]
dic_tube = dict(zip(df_tube.iloc[:, 2], df_tube.iloc[:, 0]))
dic_elbow = dict(zip(zip(df_elbow.iloc[:, 2], df_elbow.iloc[:, 3]), df_elbow.iloc[:, 0]))
dic_tee = dict(zip(zip(df_tee.iloc[:, 2], df_tee.iloc[:, 3]), df_tee.iloc[:, 0]))

# 创建处理数据按钮
tk.Button(root, text="点一下|自动出来PML命令！", command=process_data).pack(pady=20)

# 运行主窗口
root.mainloop()

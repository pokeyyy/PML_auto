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
    def generate_commands(self, indent=0, left_align=True):
        indent_str = '' if left_align else ' ' * indent
        commands = []

        commands.append(f"{indent_str}NEW {self.model_type} {self.name}")

        for key, value in self.attributes.items():
            commands.append(f"{indent_str}{key} {value}")

        for child in self.children:
            commands.extend(child.generate_commands(indent + 4, left_align))

        commands.append(f"{indent_str}END")
        return commands

#ZONE层
def create_ZONE(purp):
    return {
        "PURP": f"{purp}"
    }

#PIPE层
def create_PIPE(purp):
    return{
        # 必填项：TEMP PSPE BORE
        "PURP": f"{purp}",
        "BUIL": "false",
        "SHOP": "false",
        "TEMP": "-100000degC",  # 管道温度跟保温层厚度相关
        "PRES": "0pascal",
        "TPRESS": "0pascal",
        "PSPE": "SPECIFICATION /1C0031",
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
    }

#branch
def create_branch(hpos,tpos,hdir,tdir,hbor,tbor,hstu,purp):
    return {
        # 必填项：HPOS TPOS HDIR TDIR LHEA LTAI HBOR TBOR HCON DETA TEMP HSTU PSPE TCON
        "PURP": f"{purp}",
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
        "PSPE": "SPECIFICATION /1C0031",
        "DUTY": "'unset'",
        "DSCO": "'unset'",
        "PTSP": "'unset'",
        "INSC": "'unset'",
        "PTNB": "0",
        "PLANU": "unset",
        "DELDSG": "FALSE"
                  "\n"
    }

# elbow tee
def create_component_1(POS,ORI,SPRE,LSTU,arrive = 1,leave = 2):
    return {
        "POS": f"E {POS[0]:.3f}mm N {POS[1]:.3f}mm U {POS[2]:.3f}mm",  # 必填数据——弯头元件的坐标值
        "ORI": f"X is E{round(ORI[0])}N{round(ORI[1])}U and Z is E{round(ORI[2])}N{round(ORI[3])}U",
        # 必填数据——弯头元件的朝向
        "BUIl": "false",
        "SHOP": "true",
        "SPRE": f"SPCOMPONENT {SPRE}",  # 必填数据——管道等级中对应弯头元件的name
        "LSTU": f"SPCOMPONENT {LSTU}",  # 必填数据——管道等级中对应管道的元件name
        "ORIF": "true",
        "POSF": "true",
        "ARRIVE": f"{arrive}",
        "LEAVE": f"{leave}"
                 "\n"
    }

#redu valv
def create_component_2(POS,ORI,SPRE,LSTU,arrive = 1,leave = 2):
    return {
        "POS": f"E {POS[0]}mm N {POS[1]}mm U {POS[2]}mm",  # 必填数据——三通元件的坐标值
        "ORI": f"X is E{round(ORI[0])}N{round(ORI[1])}U",
        # 必填数据——弯头元件的朝向
        "BUIl": "false",
        "SHOP": "true",
        "SPRE": f"SPCOMPONENT {SPRE}",  # 必填数据——管道等级中对应弯头元件的name
        "LSTU": f"SPCOMPONENT {LSTU}",  # 必填数据——管道等级中对应管道的元件name
        "ORIF": "true",
        "POSF": "true",
        "ARRIVE": f"{arrive}",
        "LEAVE": f"{leave}"
                "\n"
    }

#flan gask
def create_component_3(POS,ORI,SPRE,LSTU,arrive,leave):
    return {
        "POS": f"E {POS[0]}mm N {POS[1]}mm U {POS[2]}mm",  # 必填数据——三通元件的坐标值
        "ORI": f"X is E{round(ORI[0])}N{round(ORI[1])}U",
        # 必填数据——弯头元件的朝向
        "BUIl": "false",
        "SHOP": "true",
        "SPRE": f"SPCOMPONENT {SPRE}",  # 必填数据——管道等级中对应弯头元件的name
        "LSTU": f"SPCOMPONENT {LSTU}",  # 必填数据——管道等级中对应管道的元件name
        "ORIF": "true",
        "POSF": "true",
        "ARRIVE": f"{arrive}",
        "LEAVE": f"{leave}"
                "\n"
    }
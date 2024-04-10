#coding=UTF-8
# ----------------------------------------------
# Script Recorded by Ansys Electronics Desktop Version 2021.2.0
# 用于利用HFSS脚本构建叠层
# ----------------------------------------------
#变量定义
Project_Name='Project1'
Design_Name=Project_Name
L_Sub="5mm"#板子长度
W_Sub="5mm"#板子宽度
Num_Stack=8#叠层层数 必须是2的倍数 最小是双层板
H_M="0.035mm"#铜厚
H_Core="0.254mm"#芯板厚度
H_PP="0.16mm"#PP厚度
names = locals()
for i in range(1, Num_Stack+1):
    names['H_M%s' % i] = H_M
for i in range(1, Num_Stack+1,2):
    names['H_Core_M%s_To_M%s' % (i,i+1)] = H_Core
for i in range(2, Num_Stack,2):
    names['H_PP_M%s_To_M%s' % (i,i+1)] = H_PP


#HFSS脚本实现
import ScriptEnv
#ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
#oDesktop.RestoreWindow()
oProject = oDesktop.SetActiveProject(Project_Name)
oProject.InsertDesign("HFSS", Design_Name, "HFSS Modal Network", "")
oDesign = oProject.SetActiveDesign(Design_Name)
#设置HFSS变量
oDesign.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:LocalVariableTab",
            [
                "NAME:PropServers", 
                "LocalVariables"
            ],
            [
                "NAME:NewProps",
                [
                    "NAME:L_Sub",
                    "PropType:="		, "VariableProp",
                    "UserDef:="		, True,
                    "Value:="		, L_Sub
                ]
            ]
        ]
    ])
oDesign.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:LocalVariableTab",
            [
                "NAME:PropServers", 
                "LocalVariables"
            ],
            [
                "NAME:NewProps",
                [
                    "NAME:W_Sub",
                    "PropType:="		, "VariableProp",
                    "UserDef:="		, True,
                    "Value:="		, W_Sub
                ]
            ]
        ]
    ])
for i in range(1,Num_Stack+1):
    oDesign.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:LocalVariableTab",
            [
                "NAME:PropServers", 
                "LocalVariables"
            ],
            [
                "NAME:NewProps",
                [
                    "NAME:H_M%s"%(i),
                    "PropType:="		, "VariableProp",
                    "UserDef:="		, True,
                    "Value:="		, eval('H_M%s'%(i))
                ]
            ]
        ]
    ])

for i in range(1,Num_Stack+1,2):
     oDesign.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:LocalVariableTab",
            [
                "NAME:PropServers", 
                "LocalVariables"
            ],
            [
                "NAME:NewProps",
                [
                    "NAME:H_Core_M%s_To_M%s"%(i,i+1),
                    "PropType:="		, "VariableProp",
                    "UserDef:="		, True,
                    "Value:="		, eval('H_Core_M%s_To_M%s'%(i,i+1))
                ]
            ]
        ]
    ])
for i in range(2,Num_Stack,2):
     oDesign.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:LocalVariableTab",
            [
                "NAME:PropServers", 
                "LocalVariables"
            ],
            [
                "NAME:NewProps",
                [
                    "NAME:H_PP_M%s_To_M%s"%(i,i+1),
                    "PropType:="		, "VariableProp",
                    "UserDef:="		, True,
                    "Value:="		, eval('H_PP_M%s_To_M%s'%(i,i+1))
                ]
            ]
        ]
    ])
#画模型
oEditor = oDesign.SetActiveEditor("3D Modeler")
#画金属层
Z_First='0mm'#用于表示创建模型每次的高度
for i in range(1,Num_Stack+1):
    oEditor.CreateBox(
    [
        "NAME:BoxParameters",
        "XPosition:="		, "-L_Sub/2",
        "YPosition:="		, "-W_Sub/2",
        "ZPosition:="		, Z_First,
        "XSize:="		, "L_Sub",
        "YSize:="		, "W_Sub",
        "ZSize:="		,'H_M%s'%(i)
    ], 
    [
        "NAME:Attributes",
        "Name:="		, "M%s"%(i),
        "Flags:="		, "",
        "Color:="		, "(255 128 0)",
        "Transparency:="	, 0.6,
        "PartCoordinateSystem:=", "Global",
        "UDMId:="		, "",
        "MaterialValue:="	, "\"copper\"",
        "SurfaceMaterialValue:=", "\"\"",
        "SolveInside:="		, True,
        "ShellElement:="	, False,
        "ShellElementThickness:=", "0mm",
        "IsMaterialEditable:="	, True,
        "UseMaterialAppearance:=", False,
        "IsLightweight:="	, False
    ])
    Z_First=Z_First+'+H_M%s'%(i)
    if i%2!=0:
        Z_First=Z_First+'+H_Core_M%s_To_M%s'%(i,i+1)
    else:
        Z_First=Z_First+'+H_PP_M%s_To_M%s'%(i,i+1)
        
#画金属空隙中的PP
Z_First='H_M1+H_Core_M1_To_M2'#用于表示创建模型每次的高度
for i in range(2,Num_Stack):
    oEditor.CreateBox(
    [
        "NAME:BoxParameters",
        "XPosition:="		, "-L_Sub/2",
        "YPosition:="		, "-W_Sub/2",
        "ZPosition:="		, Z_First,
        "XSize:="		, "L_Sub",
        "YSize:="		, "W_Sub",
        "ZSize:="		, 'H_M%s'%(i)
    ], 
    [
        "NAME:Attributes",
        "Name:="		, "PP_M%s"%(i),
        "Flags:="		, "",
        "Color:="		, "(255 128 0)",
        "Transparency:="	, 1,
        "PartCoordinateSystem:=", "Global",
        "UDMId:="		, "",
        "MaterialValue:="	, "\"vacuum\"",
        "SurfaceMaterialValue:=", "\"\"",
        "SolveInside:="		, True,
        "ShellElement:="	, False,
        "ShellElementThickness:=", "0mm",
        "IsMaterialEditable:="	, True,
        "UseMaterialAppearance:=", False,
        "IsLightweight:="	, False
    ])
    Z_First=Z_First+'+H_M%s'%(i)
    if i%2!=0:
        Z_First=Z_First+'+H_Core_M%s_To_M%s'%(i,i+1)
    else:
        Z_First=Z_First+'+H_PP_M%s_To_M%s'%(i,i+1)

#画Core层
Z_First='H_M1'
for i in range(1,Num_Stack+1,2):
    oEditor.CreateBox(
    [
        "NAME:BoxParameters",
        "XPosition:="		, "-L_Sub/2",
        "YPosition:="		, "-W_Sub/2",
        "ZPosition:="		, Z_First,
        "XSize:="		, "L_Sub",
        "YSize:="		, "W_Sub",
        "ZSize:="		, 'H_Core_M%s_To_M%s'%(i,i+1)
    ], 
    [
        "NAME:Attributes",
        "Name:="		, "Core_M%s_To_M%s"%(i,i+1),
        "Flags:="		, "",
        "Color:="		, "(255 128 0)",
        "Transparency:="	, 1,
        "PartCoordinateSystem:=", "Global",
        "UDMId:="		, "",
        "MaterialValue:="	, "\"vacuum\"",
        "SurfaceMaterialValue:=", "\"\"",
        "SolveInside:="		, True,
        "ShellElement:="	, False,
        "ShellElementThickness:=", "0mm",
        "IsMaterialEditable:="	, True,
        "UseMaterialAppearance:=", False,
        "IsLightweight:="	, False
    ])
    Z_First=Z_First+'+H_Core_M%s_To_M%s'%(i,i+1)+'+H_M%s'%(i+1)+'+H_PP_M%s_To_M%s'%(i+1,i+2)+'+H_M%s'%(i+2)
#画PP层
Z_First='H_M1+H_Core_M1_To_M2+H_M2'
for i in range(2,Num_Stack,2):
    oEditor.CreateBox(
    [
        "NAME:BoxParameters",
        "XPosition:="		, "-L_Sub/2",
        "YPosition:="		, "-W_Sub/2",
        "ZPosition:="		, Z_First,
        "XSize:="		, "L_Sub",
        "YSize:="		, "W_Sub",
        "ZSize:="		, 'H_PP_M%s_To_M%s'%(i,i+1)
    ], 
    [
        "NAME:Attributes",
        "Name:="		, "PP_M%s_To_M%s"%(i,i+1),
        "Flags:="		, "",
        "Color:="		, "(255 128 0)",
        "Transparency:="	, 1,
        "PartCoordinateSystem:=", "Global",
        "UDMId:="		, "",
        "MaterialValue:="	, "\"vacuum\"",
        "SurfaceMaterialValue:=", "\"\"",
        "SolveInside:="		, True,
        "ShellElement:="	, False,
        "ShellElementThickness:=", "0mm",
        "IsMaterialEditable:="	, True,
        "UseMaterialAppearance:=", False,
        "IsLightweight:="	, False
    ])
    Z_First=Z_First+'+H_PP_M%s_To_M%s'%(i,i+1)+'+H_M%s'%(i+1)+'+H_Core_M%s_To_M%s'%(i+1,i+2)+'+H_M%s'%(i+2)   

#oProject.Save()


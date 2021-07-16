Imports System.Math
Imports MySql.Data.MySqlClient
Imports SldWorks
Imports SwConst
Public Class EQ_214_2431

    Public Shared A1, A2, A3, A4, A5, A6, A7, A8, A9, A10 As SldWorks.SketchArc
    Public Shared L1, L2, L3, L4, L5, L6, L7, L8, L9, L10, L11, L12, L13, L14, L15, L01, L02, L03, L04, L05 As SldWorks.SketchLine
    Public Shared DL1, DL2, DL3, DL4, DL5, DL6 As SldWorks.SketchLine
    Public Shared A1Segment, A2Segment, A3Segment, A4Segment, A5Segment, L1Segment, L2Segment, L3Segment, L4Segment, L5Segment, L6Segment, L7Segment,
                L8Segment, L9Segment, L10Segment, L11Segment, L12Segment, L13Segment, L14Segment, L15Segment, L01Segment, L02Segment, L03Segment, L04Segment, L05Segment, DL1Segment, DL2Segment,
                DL3Segment, DL4Segment, DL5Segment, DL6Segment, SketchSegment As SldWorks.SketchSegment
    Public Shared SketchSegments(), points() As Object
    Public Shared [Boolean] As Boolean
    Public Shared P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, DP1, DP2 As SldWorks.SketchPoint
    Public Class Stator_pressing_ring
        Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
        Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
        Dim NewDocument As SldWorks.ModelDoc2 = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_part.prtdot", 0, 0, 0)
        Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc '获取当前活动的文档对象
        Dim SketchManager As SldWorks.SketchManager = part.SketchManager '草图管理器
        Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager '声明变量获得特征管理器对象,用其接受特征管理器引用
        Dim Dimension As SldWorks.Dimension '模型的尺寸和公差
        Dim DisplayDimension As SldWorks.DisplayDimension '表示显示的尺寸实例
        Dim DimensionTolerance As DimensionTolerance
        Dim sketch As SldWorks.Sketch
        Dim Feature As SldWorks.Feature
        Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager '选择管理器的接口
        Const Pi As Double = 3.1415926535897931
        Public Sub Circle_table(outer_diameter#, thick#, outer_circle_table_tol_type#, outer_circle_table_tol_max$, outer_circle_table_tol_min$)
            '圆凸台(外径,厚,外圆公差类型,外圆上偏差,外圆下偏差)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            A1Segment = part.SketchManager.CreateCircleByRadius(0, 0, 0, outer_diameter / 2) '用半径画圆
            DisplayDimension = part.AddDimension2(outer_diameter * 2 / 3 * Cos(60 * Pi / 180), outer_diameter * 2 / 3 * Sin(60 * Pi / 180), 0) '标注尺寸
            DisplayDimension.SetBrokenLeader2(False, 2) '直径标注类型,水平
            If Not outer_circle_table_tol_type = 0 Then
                part.EditDimensionProperties2(0, 0, 0, "", "", False, 2, 2, True, 12, 12, "<MOD-DIAM>", "", True, "", "", False)
                SetTolvalue(outer_circle_table_tol_type, outer_circle_table_tol_max, outer_circle_table_tol_min) '公差类型,孔,轴
                DisplayDimension.ShowTolParenthesis = True '公差括号
            End If

            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, thick, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0) '拉伸,单向,
            '不合并实体,特征影响所有实体,选择特征影响的实体,从基准面开始拉伸


        End Sub
        Public Sub Circle_cut(resection_diameter#, resection_circle_tol_type#, resection_circle_tol_max$, resection_circle_tol_min$)
            '圆切除(切除圆直径,切除圆公差类型,切除圆上偏差,切除圆下偏差)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            part.SketchManager.CreateCircleByRadius(0, 0, 0, resection_diameter / 2) '用半径画圆
            DisplayDimension = part.AddDimension2(resection_diameter * 2 / 3 * Cos(45 * Pi / 180), resection_diameter * 2 / 3 * Sin(45 * Pi / 180), 0) '标注尺寸
            DisplayDimension.SetBrokenLeader2(False, 2) '直径标注类型,水平
            If Not resection_circle_tol_type = 0 Then '标注公差
                part.EditDimensionProperties2(0, 0, 0, "", "", False, 2, 2, True, 12, 12, "<MOD-DIAM>", "", True, "", "", False)
                SetTolvalue(resection_circle_tol_type, resection_circle_tol_max, resection_circle_tol_min) '公差类型,孔,轴
                DisplayDimension.ShowTolParenthesis = True '公差括号
            End If
            sketch = part.SketchManager.ActiveSketch
            part.SketchManager.InsertSketch(True) '结束草图编辑
            part.FeatureCut(True, False, True, 1, 0, 0, 0, False, False, 0, 0, 0, 0, 0, 0) '反向贯穿切除,单向
        End Sub
        Public Function ChangeDimensionVaule(vaule#) '修改标注尺寸
            Dimension = part.Parameter(DisplayDimension.GetNameForSelection())
            Dimension.SystemValue = vaule
        End Function
        Public Sub SetTolvalue(type#, max$, min$) '公差标注
            Dimension = DisplayDimension.GetDimension2(0)
            DimensionTolerance = Dimension.Tolerance
            DimensionTolerance.Type = type
            DimensionTolerance.SetFitValues(max, min)

        End Sub
        Public Function 草图方法(类型$)
            If 类型 = "固定" Then
                part.SketchAddConstraints("sgFIXED")
            ElseIf 类型 = "水平" Then
                part.SketchAddConstraints("sgHORIZONTAL2D")
            ElseIf 类型 = "点水平排布" Then
                part.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            ElseIf 类型 = "点竖直排布" Then
                part.SketchAddConstraints("sgVERTICALPOINTS2D")
            ElseIf 类型 = "垂直" Then
                part.SketchAddConstraints("sgVERTICAL2D")
            ElseIf 类型 = "共线" Then
                part.SketchAddConstraints("sgCOLINEAR")
            ElseIf 类型 = "全等" Then
                part.SketchAddConstraints("sgCORADIAL")
            ElseIf 类型 = "相互垂直" Then
                part.SketchAddConstraints("sgPERPENDICULAR")
            ElseIf 类型 = "平行" Then
                part.SketchAddConstraints("sgPARALLEL")
            ElseIf 类型 = "相切" Then
                part.SketchAddConstraints("sgTANGENT")
            ElseIf 类型 = "同心" Then
                part.SketchAddConstraints("sgCONCENTRIC")
            ElseIf 类型 = "重合" Then
                part.SketchAddConstraints("sgCOINCIDENT")
            ElseIf 类型 = "对称" Then
                part.SketchAddConstraints("sgSYMMETRIC")
            ElseIf 类型 = "相等" Then
                part.SketchAddConstraints("sgSAMELENGTH")
            ElseIf 类型 = "合并" Then
                part.SketchAddConstraints("sgMERGEPOINTS")
            ElseIf 类型 = "相等曲率" Then
                part.SketchAddConstraints("sgEQUALCURV3DALIGN")
            ElseIf 类型 = "曲线长度相等" Then
                part.SketchAddConstraints("sgSAMECURVELENGTH")
            ElseIf 类型 = "镜像" Then
                part.SketchMirror()
            End If

        End Function
        Public Function 基准轴(基准$) As SldWorks.Feature
            Dim AXIS As SldWorks.Feature
            Dim line1 As SldWorks.SketchLine
            Dim line1Segment As SldWorks.SketchSegment
            If 基准 = "Z" Then
                part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, -0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "Y" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "X" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0.01, 0, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            基准轴 = AXIS
        End Function
        Public Function Initial_setting(type#) '初始设定
            If type = 0 Then '关闭捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 1 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 2 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 3 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            End If

        End Function
        Public Function 直径标注(类型$)
            If 类型 = "水平" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderHorizontalText) '直径水平
            ElseIf 类型 = "跟随" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swSolidLeaderAlignedText) '直径跟随标注线
            ElseIf 类型 = "断开" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderAlignedText) '标注与线断开
            End If
        End Function
        Public Sub [End](dz$)
            part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
            part.SaveAs3(dz, 0, 8)
        End Sub
        Public Sub set_attribute(name$, code$, material$)

            Dim cusproper As SldWorks.CustomPropertyManager
            cusproper = part.Extension.CustomPropertyManager("")
            cusproper.Set2("名称", name)
            cusproper.Set2("代号", code)
            cusproper.Set2("材料", material)
        End Sub
    End Class
    Public Class Balance_ring
        Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
        Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
        Dim NewDocument As SldWorks.ModelDoc2 = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_part.prtdot", 0, 0, 0)
        Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc '获取当前活动的文档对象
        Dim SketchManager As SldWorks.SketchManager = part.SketchManager '草图管理器
        Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager '声明变量获得特征管理器对象,用其接受特征管理器引用
        Dim Dimension As SldWorks.Dimension '模型的尺寸和公差
        Dim DisplayDimension As SldWorks.DisplayDimension '表示显示的尺寸实例
        Dim DimensionTolerance As DimensionTolerance
        Dim sketch As SldWorks.Sketch
        Dim Feature As SldWorks.Feature
        Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager '选择管理器的接口
        Const Pi As Double = 3.1415926535897931
        Public Sub Circle_table(outer_diameter#, thick#)
            '圆凸台（外径,厚）

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            part.SketchManager.CreateCircleByRadius(0, 0, 0, outer_diameter / 2) '用半径画圆
            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, thick, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0) '拉伸,单向,
            '不合并实体, 特征影响所有实体, 选择特征影响的实体, 从基准面开始拉伸

            part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '绘制右视尺寸标注
            part.SketchManager.InsertSketch(True) '插入新的草图
            P1 = part.SketchManager.CreatePoint(0, -outer_diameter / 2, 0)
            P2 = part.SketchManager.CreatePoint(0, outer_diameter / 2, 0)
            P1.Select4(False, Nothing)
            P2.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(0, 0, 7 * thick) '标注尺寸
            part.EditDimensionProperties2(0, 0, 0, "", "", True, 9, 2, True, 12, 12, "<MOD-DIAM>", "", True, "", "", False) '修改尺寸标注内容,添加直径符号
            part.SketchManager.InsertSketch(True)
        End Sub

        Public Sub Circle_cut(resection_diameter#)
            '圆切除(切除直径)
            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            part.SketchManager.CreateCircleByRadius(0, 0, 0, resection_diameter / 2) '用半径画圆
            DisplayDimension = part.AddDimension2(resection_diameter * 3 / 4 * Cos(30 * Pi / 180), resection_diameter * 3 / 4 * Sin(30 * Pi / 180), 0) '标注尺寸
            DisplayDimension.SetBrokenLeader2(False, 2) '直径标注类型,水平
            sketch = part.SketchManager.ActiveSketch
            part.SketchManager.InsertSketch(True) '结束草图编辑
            part.FeatureCut(True, False, True, 1, 0, 0, 0, False, False, 0, 0, 0, 0, 0, 0) '反向贯穿切除,单向
        End Sub
        Public Sub Positioning_hole(positioning_diameter#, positioning_hole_diameter#, positioning_hole_angle#, positioning_hole_num#, positioning_hole_gtol_type$, positioning_hole_gtol$)
            '定位孔(定位直径,定位孔直径,定位孔角度,定位孔数量,定位孔形位公差类型,定位孔形位公差)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            A1Segment = part.SketchManager.CreateCircleByRadius(0, -positioning_diameter / 2, 0, positioning_hole_diameter / 2) '用半径画圆
            A1Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(-positioning_diameter * 2 / 3 * Cos(70 * Pi / 180), -positioning_diameter * 2 / 3 * Sin(70 * Pi / 180), 0) '标注尺寸
            DisplayDimension.SetBrokenLeader2(False, 2) '直径标注类型,水平
            part.EditDimensionProperties2(0, 0, 0, "", "", True, 9, 2, True, 12, 12, positioning_hole_num.ToString + "x" + "<MOD-DIAM>", "", True, "", "", False) '修改尺寸内容

            part.Extension.SelectByID2(A1Segment.GetName, "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
            Dim myGtol As Object '标注形位公差
            Dim myAnno As Object
            myGtol = part.InsertGtol()
            myGtol.SetFrameSymbols2(1, positioning_hole_gtol_type, True, "", False, "", "", "", "")
            myGtol.SetFrameValues(1, positioning_hole_gtol, "", "", "", "")
            myGtol.SetFrameSymbols2(2, "", False, "", False, "", "", "", "")
            myGtol.SetFrameValues(2, "", "", "", "", "")
            myGtol.SetPTZHeight("", False)
            myGtol.SetCompositeFrame(False)
            myGtol.SetText(4, "")
            myGtol.SetBetweenTwoPoints(False, "", "")
            myAnno = myGtol.GetAnnotation()
            myAnno.SetPosition(-positioning_diameter * 2 / 3 * Cos(60 * Pi / 180) - 0.024, -positioning_diameter * 3 / 4 * Sin(80 * Pi / 180), 0) '形位公差位置
            myAnno.SetLeader3(2, 0, True, False, False, False) '设置引线

            A1Segment.Select4(False, Nothing)
            part.Extension.RotateOrCopy(True, positioning_hole_num - 1, False, 0, 0, 0, 0, 0, 1, -Pi * positioning_hole_angle / 180) '旋转、缩放和复制草图
            part.SketchManager.InsertSketch(True)
            part.FeatureCut(True, False, True, 1, 0, 0, 0, False, False, 0, 0, 0, 0, 0, 0) '反向贯穿切除,单向

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '绘制构造圆尺寸标注
            part.SketchManager.InsertSketch(True) '插入新的草图
            A2Segment = part.SketchManager.CreateCircleByRadius(0, 0, 0, positioning_diameter / 2) '用半径画圆
            part.SketchManager.CreateConstructionGeometry() '设置为辅助线
            A2Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(positioning_diameter * 2 / 3 * Cos(75 * Pi / 180), positioning_diameter * 2 / 3 * Sin(75 * Pi / 180), 0) '标注尺寸
            DisplayDimension.SetBrokenLeader2(False, 2) '直径标注类型,水平
            part.EditDimensionProperties2(1, 0, 0, "", "", True, 9, 2, True, 12, 12, "<MOD-DIAM>", "", True, "", "", False) '设置公差为基准
            part.SketchManager.InsertSketch(True)
        End Sub
        Public Sub Array_hole(array_diameter#, array_hole_diameter#, array_hole_angle#, array_hole_num#, positioning_hole_num#, array_hole_gtol_type$, array_hole_gtol$)
            '阵列孔(阵列直径,阵列孔直径,阵列孔角度,阵列孔数量,定位孔数量,阵列孔形位公差类型,阵列孔形位公差)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            A1 = part.SketchManager.CreateCircleByRadius(0, -array_diameter / 2, 0, array_hole_diameter / 2) '用半径画圆
            A1Segment = A1
            A1Segment.Select4(False, Nothing)
            part.Extension.RotateOrCopy(False, 1, False, 0, 0, 0, 0, 0, 1, Pi * array_hole_angle / 180) '旋转圆
            A1Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(array_diameter * 2 / 3 * Cos(60 * Pi / 180), -array_diameter * 2 / 3 * Sin(60 * Pi / 180), 0) '标注尺寸
            DisplayDimension.SetBrokenLeader2(False, 2) '直径标注类型,水平
            part.EditDimensionProperties2(0, 0, 0, "", "", True, 9, 2, True, 12, 12, array_hole_num.ToString + "x" + "<MOD-DIAM>", "", True, "", "", False)

            part.Extension.SelectByID2(A1Segment.GetName, "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
            Dim myGtol As Object '标注形位公差
            Dim myAnno As Object
            myGtol = part.InsertGtol()
            myGtol.SetFrameSymbols2(1, array_hole_gtol_type, True, "", False, "", "", "", "")
            myGtol.SetFrameValues(1, array_hole_gtol, "", "", "", "")
            myGtol.SetFrameSymbols2(2, "", False, "", False, "", "", "", "")
            myGtol.SetFrameValues(2, "", "", "", "", "")
            myGtol.SetPTZHeight("", False)
            myGtol.SetCompositeFrame(False)
            myGtol.SetText(4, "")
            myGtol.SetBetweenTwoPoints(False, "", "")
            myAnno = myGtol.GetAnnotation()
            myAnno.SetPosition(array_diameter * 2 / 3 * Cos(60 * Pi / 180), -array_diameter * 2 / 3 * Sin(80 * Pi / 180), 0)
            myAnno.SetLeader3(2, 0, True, False, False, False)

            A1Segment.Select4(False, Nothing)
            part.Extension.RotateOrCopy(True, array_hole_num / positioning_hole_num - 1, False, 0, 0, 0, 0, 0, 1, Pi * array_hole_angle / 180) '旋转复制
            part.SketchManager.InsertSketch(True)
            part.FeatureCut(True, False, True, 1, 0, 0, 0, False, False, 0, 0, 0, 0, 0, 0) '反向贯穿切除,单向
            Feature = SelectionMgr.GetSelectedObject6(1, -1) '获得所选择的对象引用,无论标记如何,所有选择
            Dim 基准轴Z As SldWorks.Feature
            基准轴Z = 基准轴("Z")
            基准轴Z.Select2(False, 1) '选择并标记
            part.Extension.SelectByID2(Feature.Name, "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0) '选择特征
            part.FeatureManager.FeatureCircularPattern5(positioning_hole_num, -Pi * array_hole_angle * (array_hole_num / positioning_hole_num + 1) / 180, False, "NULL", False, False, False, False, False, False, 1, 0, "NULL", False)
            基准轴Z.Select2(False, Nothing) '基准轴与阵列特征相关联不能删除；特征循环模式(阵列)
            part.BlankRefGeom() '隐藏


            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            A2Segment = part.SketchManager.CreateCircleByRadius(0, 0, 0, array_diameter / 2) '用半径画圆
            part.SketchManager.CreateConstructionGeometry() '设置为辅助线
            A2Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(array_diameter * 2 / 3 * Cos(45 * Pi / 180), array_diameter * 2 / 3 * Sin(45 * Pi / 180), 0) '标注尺寸
            DisplayDimension.SetBrokenLeader2(False, 2) '直径标注类型,水平
            part.EditDimensionProperties2(1, 0, 0, "", "", True, 9, 2, True, 12, 12, "<MOD-DIAM>", "", True, "", "", False) '设置公差为基本

            L01Segment = part.SketchManager.CreateCenterLine(0, array_diameter * 1 / 3, 0, 0, array_diameter * 3 / 4, 0) '绘制角度标注
            part.Extension.RotateOrCopy(False, 1, False, 0, 0, 0, 0, 0, 1, Pi * 2 * array_hole_angle / 180) '旋转
            L02Segment = part.SketchManager.CreateCenterLine(0, array_diameter * 1 / 3, 0, 0, array_diameter * 3 / 4, 0)
            part.Extension.RotateOrCopy(False, 1, False, 0, 0, 0, 0, 0, 1, Pi * array_hole_angle / 180) '旋转
            L03Segment = part.SketchManager.CreateCenterLine(0, array_diameter * 1 / 3, 0, 0, array_diameter * 3 / 4, 0)
            L01Segment.Select4(False, Nothing)
            L02Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(-array_diameter * 3 / 4 * Cos((90 - 5 / 2 * array_hole_angle) * Pi / 180), array_diameter * 3 / 4 * Sin((90 - 5 / 2 * array_hole_angle) * Pi / 180), 0) '标注尺寸
            L03Segment.Select4(False, Nothing)
            L02Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(array_diameter * 3 / 4 * Cos((90 - array_hole_angle / 2) * Pi / 180), array_diameter * 3 / 4 * Sin((90 - array_hole_angle / 2) * Pi / 180), 0) '标注尺寸
            part.SketchManager.InsertSketch(True)
        End Sub
        Public Function ChangeDimensionVaule(vaule#) '修改标注尺寸
            Dimension = part.Parameter(DisplayDimension.GetNameForSelection())
            Dimension.SystemValue = vaule
        End Function
        Public Sub SetTolvalue(type#, max$, min$) '公差标注
            Dimension = DisplayDimension.GetDimension2(0)
            DimensionTolerance = Dimension.Tolerance
            DimensionTolerance.Type = type
            DimensionTolerance.SetFitValues(max, min)

        End Sub
        Public Function 草图方法(类型$)
            If 类型 = "固定" Then
                part.SketchAddConstraints("sgFIXED")
            ElseIf 类型 = "水平" Then
                part.SketchAddConstraints("sgHORIZONTAL2D")
            ElseIf 类型 = "点水平排布" Then
                part.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            ElseIf 类型 = "点竖直排布" Then
                part.SketchAddConstraints("sgVERTICALPOINTS2D")
            ElseIf 类型 = "垂直" Then
                part.SketchAddConstraints("sgVERTICAL2D")
            ElseIf 类型 = "共线" Then
                part.SketchAddConstraints("sgCOLINEAR")
            ElseIf 类型 = "全等" Then
                part.SketchAddConstraints("sgCORADIAL")
            ElseIf 类型 = "相互垂直" Then
                part.SketchAddConstraints("sgPERPENDICULAR")
            ElseIf 类型 = "平行" Then
                part.SketchAddConstraints("sgPARALLEL")
            ElseIf 类型 = "相切" Then
                part.SketchAddConstraints("sgTANGENT")
            ElseIf 类型 = "同心" Then
                part.SketchAddConstraints("sgCONCENTRIC")
            ElseIf 类型 = "重合" Then
                part.SketchAddConstraints("sgCOINCIDENT")
            ElseIf 类型 = "对称" Then
                part.SketchAddConstraints("sgSYMMETRIC")
            ElseIf 类型 = "相等" Then
                part.SketchAddConstraints("sgSAMELENGTH")
            ElseIf 类型 = "合并" Then
                part.SketchAddConstraints("sgMERGEPOINTS")
            ElseIf 类型 = "相等曲率" Then
                part.SketchAddConstraints("sgEQUALCURV3DALIGN")
            ElseIf 类型 = "曲线长度相等" Then
                part.SketchAddConstraints("sgSAMECURVELENGTH")
            ElseIf 类型 = "镜像" Then
                part.SketchMirror()
            End If

        End Function
        Public Function 基准轴(基准$) As SldWorks.Feature
            Dim AXIS As SldWorks.Feature
            Dim line1 As SldWorks.SketchLine
            Dim line1Segment As SldWorks.SketchSegment
            If 基准 = "Z" Then
                part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, -0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "Y" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "X" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0.01, 0, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            基准轴 = AXIS
        End Function
        Public Function Initial_setting(type#)
            If type = 0 Then '关闭捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 1 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 2 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 3 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            End If

        End Function
        Public Function 直径标注(类型$)
            If 类型 = "水平" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderHorizontalText) '直径水平
            ElseIf 类型 = "跟随" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swSolidLeaderAlignedText) '直径跟随标注线
            ElseIf 类型 = "断开" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderAlignedText) '标注与线断开
            End If
        End Function
        Public Sub [End](dz$)
            part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
            part.SaveAs3(dz, 0, 8)
        End Sub
        Public Sub set_attribute(name$, code$, material$)

            Dim cusproper As SldWorks.CustomPropertyManager
            cusproper = part.Extension.CustomPropertyManager("")
            cusproper.Set2("名称", name)
            cusproper.Set2("代号", code)
            cusproper.Set2("材料", material)
        End Sub
    End Class
    Public Class Flange
        Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
        Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
        Dim NewDocument As SldWorks.ModelDoc2 = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_part.prtdot", 0, 0, 0)
        Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc '获取当前活动的文档对象
        Dim SketchManager As SldWorks.SketchManager = part.SketchManager '草图管理器
        Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager '声明变量获得特征管理器对象,用其接受特征管理器引用
        Dim Dimension As SldWorks.Dimension '模型的尺寸和公差
        Dim DisplayDimension As SldWorks.DisplayDimension '表示显示的尺寸实例
        Dim DimensionTolerance As DimensionTolerance
        Dim sketch As SldWorks.Sketch
        Dim Feature As SldWorks.Feature
        Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager '选择管理器的接口
        Const Pi As Double = 3.1415926535897931
        Public Sub Circle_table(outer_diameter#, thick#)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            A1Segment = part.SketchManager.CreateCircleByRadius(0, 0, 0, outer_diameter / 2) '用半径画圆
            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, thick, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0) '拉伸,单向,
            '不合并实体,特征影响所有实体,选择特征影响的实体,从基准面开始拉伸


        End Sub
        Public Sub Circle_cut(resection_diameter#)
            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            part.SketchManager.CreateCircleByRadius(0, 0, 0, resection_diameter / 2) '用半径画圆
            DisplayDimension = part.AddDimension2(resection_diameter * 2 / 3 * Cos(45 * Pi / 180), resection_diameter * 2 / 3 * Sin(45 * Pi / 180), 0) '标注尺寸
            sketch = part.SketchManager.ActiveSketch
            part.SketchManager.InsertSketch(True) '结束草图编辑
            part.FeatureCut(True, False, True, 1, 0, 0, 0, False, False, 0, 0, 0, 0, 0, 0) '反向贯穿切除,单向
        End Sub
        Public Sub Circle_step(outer_diameter#, step_diameter#, step_thick#, thick#)
            '圆周台阶(外径,台阶直径,台阶厚,法兰厚)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            part.SketchManager.CreateCircleByRadius(0, 0, 0, outer_diameter / 2) '用半径画圆
            part.AddDimension2(outer_diameter / 2 + 0.02, 0, 0) '标注尺寸
            part.SketchManager.CreateCircleByRadius(0, 0, 0, step_diameter / 2) '用半径画圆
            part.AddDimension2(step_diameter / 2 + 0.02, 0, 0) '标注尺寸
            part.SketchManager.InsertSketch(True) '结束草图编辑
            part.FeatureCut(True, False, True, 0, 0, step_thick, 0, False, False, 0, 0, 0, 0, 0, 0) '切除指定厚度
            part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            P1 = part.SketchManager.CreatePoint(0, -step_diameter / 2, 0)
            P2 = part.SketchManager.CreatePoint(-step_thick, -step_diameter / 2, 0)
            P3 = part.SketchManager.CreatePoint(-thick, -outer_diameter / 2, 0)
            P1.Select4(False, Nothing)
            P2.Select4(True, Nothing)
            DisplayDimension = part.AddHorizontalDimension2(0, -outer_diameter / 2 - thick / 4, step_thick / 2)
            P1.Select4(False, Nothing)
            P3.Select4(True, Nothing)
            DisplayDimension = part.AddHorizontalDimension2(0, -outer_diameter / 2 - thick / 2, thick / 2)
            part.SketchManager.InsertSketch(True)


        End Sub
        Public Function ChangeDimensionVaule(vaule#) '修改标注尺寸
            Dimension = part.Parameter(DisplayDimension.GetNameForSelection())
            Dimension.SystemValue = vaule
        End Function
        Public Sub SetTolvalue(type#, max$, min$) '公差标注
            Dimension = DisplayDimension.GetDimension2(0)
            DimensionTolerance = Dimension.Tolerance
            DimensionTolerance.Type = type
            DimensionTolerance.SetFitValues(max, min)

        End Sub
        Public Function 草图方法(类型$)
            If 类型 = "固定" Then
                part.SketchAddConstraints("sgFIXED")
            ElseIf 类型 = "水平" Then
                part.SketchAddConstraints("sgHORIZONTAL2D")
            ElseIf 类型 = "点水平排布" Then
                part.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            ElseIf 类型 = "点竖直排布" Then
                part.SketchAddConstraints("sgVERTICALPOINTS2D")
            ElseIf 类型 = "垂直" Then
                part.SketchAddConstraints("sgVERTICAL2D")
            ElseIf 类型 = "共线" Then
                part.SketchAddConstraints("sgCOLINEAR")
            ElseIf 类型 = "全等" Then
                part.SketchAddConstraints("sgCORADIAL")
            ElseIf 类型 = "相互垂直" Then
                part.SketchAddConstraints("sgPERPENDICULAR")
            ElseIf 类型 = "平行" Then
                part.SketchAddConstraints("sgPARALLEL")
            ElseIf 类型 = "相切" Then
                part.SketchAddConstraints("sgTANGENT")
            ElseIf 类型 = "同心" Then
                part.SketchAddConstraints("sgCONCENTRIC")
            ElseIf 类型 = "重合" Then
                part.SketchAddConstraints("sgCOINCIDENT")
            ElseIf 类型 = "对称" Then
                part.SketchAddConstraints("sgSYMMETRIC")
            ElseIf 类型 = "相等" Then
                part.SketchAddConstraints("sgSAMELENGTH")
            ElseIf 类型 = "合并" Then
                part.SketchAddConstraints("sgMERGEPOINTS")
            ElseIf 类型 = "相等曲率" Then
                part.SketchAddConstraints("sgEQUALCURV3DALIGN")
            ElseIf 类型 = "曲线长度相等" Then
                part.SketchAddConstraints("sgSAMECURVELENGTH")
            ElseIf 类型 = "镜像" Then
                part.SketchMirror()
            End If

        End Function
        Public Function 基准轴(基准$) As SldWorks.Feature
            Dim AXIS As SldWorks.Feature
            Dim line1 As SldWorks.SketchLine
            Dim line1Segment As SldWorks.SketchSegment
            If 基准 = "Z" Then
                part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, -0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "Y" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "X" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0.01, 0, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            基准轴 = AXIS
        End Function
        Public Function Initial_setting(type#)
            If type = 0 Then '关闭捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 1 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 2 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 3 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            End If

        End Function
        Public Function 直径标注(类型$)
            If 类型 = "水平" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderHorizontalText) '直径水平
            ElseIf 类型 = "跟随" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swSolidLeaderAlignedText) '直径跟随标注线
            ElseIf 类型 = "断开" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderAlignedText) '标注与线断开
            End If
        End Function
        Public Sub [End](dz$)
            part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
            part.SaveAs3(dz, 0, 8)
        End Sub
        Public Sub set_attribute(name$, code$, material$)

            Dim cusproper As SldWorks.CustomPropertyManager
            cusproper = part.Extension.CustomPropertyManager("")
            cusproper.Set2("名称", name)
            cusproper.Set2("代号", code)
            cusproper.Set2("材料", material)
        End Sub
    End Class
    Public Class Clsasp

        Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
        Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
        Dim NewDocument As SldWorks.ModelDoc2 = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_part.prtdot", 0, 0, 0)
        Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc '获取当前活动的文档对象
        Dim SketchManager As SldWorks.SketchManager = part.SketchManager '草图管理器
        Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager '声明变量获得特征管理器对象,用其接受特征管理器引用
        Dim Dimension As SldWorks.Dimension '模型的尺寸和公差
        Dim DisplayDimension As SldWorks.DisplayDimension '表示显示的尺寸实例
        Dim DimensionTolerance As DimensionTolerance
        Dim sketch As SldWorks.Sketch
        Dim Feature As SldWorks.Feature
        Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager '选择管理器的接口
        Const Pi As Double = 3.1415926535897931
        Public Sub Clsasp(length#, width#, height#, thick#)
            '扣片(长,宽,高,厚)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            L1 = part.SketchManager.CreateLine(thick, 0, 0, length, 0, 0)
            L2 = part.SketchManager.CreateLine(thick, 0, 0, thick, -(height - thick), 0)
            L3 = part.SketchManager.CreateLine(0, thick, 0, length, thick, 0)
            L4 = part.SketchManager.CreateLine(0, thick, 0, 0, -(height - thick), 0)
            L5 = part.SketchManager.CreateLine(0, -(height - thick), 0, thick, -(height - thick), 0)
            L6 = part.SketchManager.CreateLine(length, 0, 0, length, thick, 0)
            L3Segment = L3
            L4Segment = L4
            L4Segment.Select4(False, Nothing)
            DisplayDimension = part.AddVerticalDimension2(-height / 2, -height / 2, 0)
            L3Segment.Select4(False, Nothing)
            DisplayDimension = part.AddVerticalDimension2(height / 2, length / 2, 0)
            P1 = L1.GetStartPoint2
            P3 = L3.GetStartPoint2
            P1.Select4(False, Nothing)
            SketchManager.CreateFillet(thick, 1)
            P3.Select4(False, Nothing)
            SketchManager.CreateFillet(2 * thick, 1)
            part.SketchManager.InsertSketch(True) '结束草图编辑
            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, width, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0)
            part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            P4 = part.SketchManager.CreatePoint(0, 0, 0)
            P5 = part.SketchManager.CreatePoint(length, 0, 0)
            P4.Select4(False, Nothing)
            P5.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(length / 2, 0, 1.5 * width)
            part.SketchManager.InsertSketch(True)
        End Sub
        Public Function ChangeDimensionVaule(vaule#) '修改标注尺寸
            Dimension = part.Parameter(DisplayDimension.GetNameForSelection())
            Dimension.SystemValue = vaule
        End Function
        Public Sub SetTolvalue(type#, max$, min$) '公差标注
            Dimension = DisplayDimension.GetDimension2(0)
            DimensionTolerance = Dimension.Tolerance
            DimensionTolerance.Type = type
            DimensionTolerance.SetFitValues(max, min)

        End Sub
        Public Function 草图方法(类型$)
            If 类型 = "固定" Then
                part.SketchAddConstraints("sgFIXED")
            ElseIf 类型 = "水平" Then
                part.SketchAddConstraints("sgHORIZONTAL2D")
            ElseIf 类型 = "点水平排布" Then
                part.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            ElseIf 类型 = "点竖直排布" Then
                part.SketchAddConstraints("sgVERTICALPOINTS2D")
            ElseIf 类型 = "垂直" Then
                part.SketchAddConstraints("sgVERTICAL2D")
            ElseIf 类型 = "共线" Then
                part.SketchAddConstraints("sgCOLINEAR")
            ElseIf 类型 = "全等" Then
                part.SketchAddConstraints("sgCORADIAL")
            ElseIf 类型 = "相互垂直" Then
                part.SketchAddConstraints("sgPERPENDICULAR")
            ElseIf 类型 = "平行" Then
                part.SketchAddConstraints("sgPARALLEL")
            ElseIf 类型 = "相切" Then
                part.SketchAddConstraints("sgTANGENT")
            ElseIf 类型 = "同心" Then
                part.SketchAddConstraints("sgCONCENTRIC")
            ElseIf 类型 = "重合" Then
                part.SketchAddConstraints("sgCOINCIDENT")
            ElseIf 类型 = "对称" Then
                part.SketchAddConstraints("sgSYMMETRIC")
            ElseIf 类型 = "相等" Then
                part.SketchAddConstraints("sgSAMELENGTH")
            ElseIf 类型 = "合并" Then
                part.SketchAddConstraints("sgMERGEPOINTS")
            ElseIf 类型 = "相等曲率" Then
                part.SketchAddConstraints("sgEQUALCURV3DALIGN")
            ElseIf 类型 = "曲线长度相等" Then
                part.SketchAddConstraints("sgSAMECURVELENGTH")
            ElseIf 类型 = "镜像" Then
                part.SketchMirror()
            End If

        End Function
        Public Function 基准轴(基准$) As SldWorks.Feature
            Dim AXIS As SldWorks.Feature
            Dim line1 As SldWorks.SketchLine
            Dim line1Segment As SldWorks.SketchSegment
            If 基准 = "Z" Then
                part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, -0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "Y" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "X" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0.01, 0, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            基准轴 = AXIS
        End Function
        Public Function Initial_setting(type#)
            If type = 0 Then '关闭捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 1 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 2 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 3 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            End If

        End Function
        Public Function 直径标注(类型$)
            If 类型 = "水平" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderHorizontalText) '直径水平
            ElseIf 类型 = "跟随" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swSolidLeaderAlignedText) '直径跟随标注线
            ElseIf 类型 = "断开" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderAlignedText) '标注与线断开
            End If
        End Function
        Public Sub [End](dz$)
            part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
            part.SaveAs3(dz, 0, 8)
        End Sub

        Public Sub set_attribute(name$, code$, material$)

            Dim cusproper As SldWorks.CustomPropertyManager
            cusproper = part.Extension.CustomPropertyManager("")
            cusproper.Set2("名称", name)
            cusproper.Set2("代号", code)
            cusproper.Set2("材料", material)
        End Sub
    End Class
    Public Class Installation_platform

        Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
        Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
        Dim NewDocument As SldWorks.ModelDoc2 = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_part.prtdot", 0, 0, 0)
        Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc '获取当前活动的文档对象
        Dim SketchManager As SldWorks.SketchManager = part.SketchManager '草图管理器
        Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager '声明变量获得特征管理器对象,用其接受特征管理器引用
        Dim Dimension As SldWorks.Dimension '模型的尺寸和公差
        Dim DisplayDimension As SldWorks.DisplayDimension '表示显示的尺寸实例
        Dim DimensionTolerance As DimensionTolerance
        Dim sketch As SldWorks.Sketch
        Dim Feature As SldWorks.Feature
        Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager '选择管理器的接口
        Dim Annotation As SldWorks.Annotation
        Const Pi As Double = 3.1415926535897931
        Public Sub Front_and_rear_panel_installation_platform(length#, width#, height#, fillet_radius#, thick#)
            '前后板安装平台(长,宽,高,圆角半径,厚)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            L1 = part.SketchManager.CreateLine(thick, 0, 0, length, 0, 0)
            L2 = part.SketchManager.CreateLine(thick, 0, 0, thick, -(height - thick), 0)
            L3 = part.SketchManager.CreateLine(0, thick, 0, length, thick, 0)
            L4 = part.SketchManager.CreateLine(0, thick, 0, 0, -(height - thick), 0)
            L5 = part.SketchManager.CreateLine(0, -(height - thick), 0, thick, -(height - thick), 0)
            L6 = part.SketchManager.CreateLine(length, 0, 0, length, thick, 0)
            L4Segment = L4
            L5Segment = L5
            L4Segment.Select4(False, Nothing)
            DisplayDimension = part.AddVerticalDimension2(-height / 2, -height / 2 + thick, 0)
            L5Segment.Select4(False, Nothing)
            DisplayDimension = part.AddVerticalDimension2(thick / 2, -3 * height / 2, 0)
            P1 = L1.GetStartPoint2
            P3 = L3.GetStartPoint2
            P1.Select4(False, Nothing)
            SketchManager.CreateFillet(fillet_radius, 1)
            P3.Select4(False, Nothing)
            SketchManager.CreateFillet(fillet_radius + thick, 1)
            part.Extension.SelectByID2("D2@草图1@F126_2055.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
            Annotation = DisplayDimension.GetAnnotation()
            Annotation.SetPosition(2 * thick, -thick / 2, 0) '移动标注
            part.ClearSelection2(True)
            part.SketchManager.InsertSketch(True) '结束草图编辑
            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, width, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0)
            part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            P4 = part.SketchManager.CreatePoint(0, 0, 0)
            P5 = part.SketchManager.CreatePoint(length, 0, 0)
            P4.Select4(False, Nothing)
            P5.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(length / 2, 0, width + thick)
            part.SketchManager.InsertSketch(True)

        End Sub
        Public Sub Outlet_board_installation_platform(length#, width#, height#, fillet_radius#, thick#)
            '出线板安装平台(长,宽,高,圆角半径,厚)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            L1 = part.SketchManager.CreateLine(thick, 0, 0, length - thick, 0, 0)
            L2 = part.SketchManager.CreateLine(thick, 0, 0, thick, -(height - thick), 0)
            L3 = part.SketchManager.CreateLine(0, thick, 0, length, thick, 0)
            L4 = part.SketchManager.CreateLine(0, thick, 0, 0, -(height - thick), 0)
            L5 = part.SketchManager.CreateLine(length - thick, 0, 0, length - thick, -(height - thick), 0)
            L6 = part.SketchManager.CreateLine(length, thick, 0, length, -(height - thick), 0)
            L7 = part.SketchManager.CreateLine(0, -(height - thick), 0, thick, -(height - thick), 0)
            L8 = part.SketchManager.CreateLine(length - thick, -(height - thick), 0, length, -(height - thick), 0)

            L6Segment = L6
            L7Segment = L7
            L6Segment.Select4(False, Nothing)
            DisplayDimension = part.AddVerticalDimension2(3 * length / 2, -height / 2, 0)
            L7Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(thick / 2, -height - thick, 0)

            P1 = L1.GetStartPoint2
            P2 = L1.GetEndPoint2
            P3 = L3.GetStartPoint2
            P4 = L3.GetEndPoint2
            P1.Select4(False, Nothing)
            SketchManager.CreateFillet(fillet_radius, 1)
            P2.Select4(False, Nothing)
            SketchManager.CreateFillet(fillet_radius, 1)
            P3.Select4(False, Nothing)
            SketchManager.CreateFillet(fillet_radius + thick, 1)
            P4.Select4(False, Nothing)
            SketchManager.CreateFillet(fillet_radius + thick, 1)
            part.Extension.SelectByID2("D3@草图1@O126_2061.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
            Annotation = DisplayDimension.GetAnnotation()
            Annotation.SetPosition(3 * thick / 2, -thick / 2, 0) '移动标注
            part.Extension.SelectByID2("D4@草图1@O126_2061.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0) '删除标注
            part.Extension.SelectByID2("D6@草图1@O126_2061.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
            part.Extension.SelectByID2("D5@草图1@O126_2061.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
            part.EditDelete()

            'Dim SF2, SF3 As Object
            'SF2 = part.Extension.InsertSurfaceFinishSymbol3(2, 0, 0, 0, 0, 0, 1, "", "", "", "", "", "", "") '添加粗糙度标识,不允许切削
            'SF2.GetAnnotation().SetPosition2(长 / 2, 厚, 0)
            'SF2.SetAngle(0)

            'SF3 = part.Extension.InsertSurfaceFinishSymbol3(2, 0, 0, 0, 0, 0, 1, "", "", "", "", "", "", "")
            'SF3.GetAnnotation().SetPosition2(长 / 2, 0, 0)
            'SF3.SetAngle(90 * Pi / 180)


            part.SketchManager.InsertSketch(True) '结束草图编辑
            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, width, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0)
            part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            P6 = part.SketchManager.CreatePoint(0, 0, 0)
            P5 = part.SketchManager.CreatePoint(length, 0, 0)
            P6.Select4(False, Nothing)
            P5.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(length / 2, 0, width + thick)
            part.SketchManager.InsertSketch(True)
        End Sub
        Public Function ChangeDimensionVaule(vaule#) '修改标注尺寸
            Dimension = part.Parameter(DisplayDimension.GetNameForSelection())
            Dimension.SystemValue = vaule
        End Function
        Public Sub SetTolvalue(type#, max$, min$) '公差标注
            Dimension = DisplayDimension.GetDimension2(0)
            DimensionTolerance = Dimension.Tolerance
            DimensionTolerance.Type = type
            DimensionTolerance.SetFitValues(max, min)

        End Sub
        Public Function 草图方法(类型$)
            If 类型 = "固定" Then
                part.SketchAddConstraints("sgFIXED")
            ElseIf 类型 = "水平" Then
                part.SketchAddConstraints("sgHORIZONTAL2D")
            ElseIf 类型 = "点水平排布" Then
                part.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            ElseIf 类型 = "点竖直排布" Then
                part.SketchAddConstraints("sgVERTICALPOINTS2D")
            ElseIf 类型 = "垂直" Then
                part.SketchAddConstraints("sgVERTICAL2D")
            ElseIf 类型 = "共线" Then
                part.SketchAddConstraints("sgCOLINEAR")
            ElseIf 类型 = "全等" Then
                part.SketchAddConstraints("sgCORADIAL")
            ElseIf 类型 = "相互垂直" Then
                part.SketchAddConstraints("sgPERPENDICULAR")
            ElseIf 类型 = "平行" Then
                part.SketchAddConstraints("sgPARALLEL")
            ElseIf 类型 = "相切" Then
                part.SketchAddConstraints("sgTANGENT")
            ElseIf 类型 = "同心" Then
                part.SketchAddConstraints("sgCONCENTRIC")
            ElseIf 类型 = "重合" Then
                part.SketchAddConstraints("sgCOINCIDENT")
            ElseIf 类型 = "对称" Then
                part.SketchAddConstraints("sgSYMMETRIC")
            ElseIf 类型 = "相等" Then
                part.SketchAddConstraints("sgSAMELENGTH")
            ElseIf 类型 = "合并" Then
                part.SketchAddConstraints("sgMERGEPOINTS")
            ElseIf 类型 = "相等曲率" Then
                part.SketchAddConstraints("sgEQUALCURV3DALIGN")
            ElseIf 类型 = "曲线长度相等" Then
                part.SketchAddConstraints("sgSAMECURVELENGTH")
            ElseIf 类型 = "镜像" Then
                part.SketchMirror()
            End If

        End Function
        Public Function 基准轴(基准$) As SldWorks.Feature
            Dim AXIS As SldWorks.Feature
            Dim line1 As SldWorks.SketchLine
            Dim line1Segment As SldWorks.SketchSegment
            If 基准 = "Z" Then
                part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, -0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "Y" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "X" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0.01, 0, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            基准轴 = AXIS
        End Function
        Public Function Initial_setting(type#)

            If type = 0 Then '关闭捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 1 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 2 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 3 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            End If

        End Function
        Public Function 直径标注(类型$)
            If 类型 = "水平" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderHorizontalText) '直径水平
            ElseIf 类型 = "跟随" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swSolidLeaderAlignedText) '直径跟随标注线
            ElseIf 类型 = "断开" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderAlignedText) '标注与线断开
            End If
        End Function
        Public Sub [End](dz$)
            part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
            part.SaveAs3(dz, 0, 8)
        End Sub
        Public Sub set_attribute(name$, code$, material$)

            Dim cusproper As SldWorks.CustomPropertyManager
            cusproper = part.Extension.CustomPropertyManager("")
            cusproper.Set2("名称", name)
            cusproper.Set2("代号", code)
            cusproper.Set2("材料", material)
        End Sub
    End Class
    Public Class Hanging_climbing

        Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
        Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
        Dim NewDocument As SldWorks.ModelDoc2 = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_part.prtdot", 0, 0, 0)
        Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc '获取当前活动的文档对象
        Dim SketchManager As SldWorks.SketchManager = part.SketchManager '草图管理器
        Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager '声明变量获得特征管理器对象,用其接受特征管理器引用
        Dim Dimension As SldWorks.Dimension '模型的尺寸和公差
        Dim DisplayDimension As SldWorks.DisplayDimension '表示显示的尺寸实例
        Dim DimensionTolerance As DimensionTolerance
        Dim sketch As SldWorks.Sketch
        Dim Feature As SldWorks.Feature
        Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager '选择管理器的接口
        Dim Annotation As SldWorks.Annotation
        Const Pi As Double = 3.1415926535897931
        Public Sub Hanging_climbing(base_board_length#, width#, hole_board_length#, fillet_radius#, thick#, slanted_board_length#, circle_edge_distance#, diameter#, slanted_angle#)
            '吊攀(底板长,宽,孔板长,圆角半径,厚,斜板长,圆边距,直径,斜角)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            L1 = part.SketchManager.CreateLine(thick, 0, 0, base_board_length - thick, 0, 0)
            L2 = part.SketchManager.CreateLine(thick, 0, 0, thick, -(hole_board_length - thick), 0)
            L3 = part.SketchManager.CreateLine(0, thick, 0, base_board_length, thick, 0)
            L4 = part.SketchManager.CreateLine(0, thick, 0, 0, -(hole_board_length - thick), 0)

            L5 = part.SketchManager.CreateLine(base_board_length - thick, -0.001, 0, base_board_length - thick, -(slanted_board_length - thick) + 0.001, 0)
            L6 = part.SketchManager.CreateLine(base_board_length, thick, 0, base_board_length, -(slanted_board_length - thick) + 0.001, 0)
            L7 = part.SketchManager.CreateLine(0, -(hole_board_length - thick), 0, thick, -(hole_board_length - thick), 0)
            L8 = part.SketchManager.CreateLine(base_board_length - thick, -(slanted_board_length - thick), 0, base_board_length, -(slanted_board_length - thick), 0)

            L1Segment = L1
            L3Segment = L3
            L4Segment = L4
            L5Segment = L5
            L6Segment = L6
            L7Segment = L7
            L8Segment = L8
            L1Segment.Select4(False, Nothing)
            L3Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(-thick, -thick, 0) '标注尺寸
            L3Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(base_board_length / 2, 2 * thick, 0) '标注尺寸
            L6Segment.Select4(False, Nothing)
            L3Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(2 * base_board_length / 3, -base_board_length / 3, 0) '标注尺寸
            ChangeDimensionVaule(slanted_angle * Pi / 180)
            'part.Extension.RotateOrCopy(False, 1, False, 长, 厚, 0, 0, 0, 1, -Pi * (角度 - 90) / 180) '旋转、缩放和复制草图
            L6Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(3 * base_board_length / 2, -slanted_board_length / 3, 0) '标注尺寸
            ChangeDimensionVaule(slanted_board_length)
            L4Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(-3 * thick, -hole_board_length / 2, 0) '标注尺寸

            P1 = L1.GetStartPoint2
            P2 = L1.GetEndPoint2
            P3 = L3.GetStartPoint2
            P4 = L3.GetEndPoint2
            P5 = L5.GetStartPoint2
            P6 = L5.GetEndPoint2
            P7 = L6.GetEndPoint2
            P8 = L8.GetStartPoint2
            P9 = L8.GetEndPoint2

            P7.Select4(False, Nothing)
            P9.Select4(True, Nothing)
            草图方法(“合并”)
            L8Segment.Select4(False, Nothing)
            L6Segment.Select4(True, Nothing)
            草图方法(“相互垂直”)
            L8Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(3 * base_board_length / 2, -slanted_board_length, 0) '标注尺寸
            ChangeDimensionVaule(thick)
            L5Segment.Select4(False, Nothing)
            L6Segment.Select4(True, Nothing)
            草图方法(“平行”)
            P8.Select4(False, Nothing)
            P6.Select4(True, Nothing)
            草图方法(“合并”)
            P2.Select4(False, Nothing)
            P5.Select4(True, Nothing)
            草图方法(“合并”)
            P1.Select4(False, Nothing)
            SketchManager.CreateFillet(fillet_radius, 1)
            P2.Select4(False, Nothing)
            SketchManager.CreateFillet(fillet_radius, 1)
            P3.Select4(False, Nothing)
            SketchManager.CreateFillet(fillet_radius + thick, 1)
            P4.Select4(False, Nothing)
            SketchManager.CreateFillet(fillet_radius + thick, 1)

            L7Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(thick / 2, -hole_board_length - thick, 0) '标注尺寸
            L5Segment.Select4(False, Nothing)
            L1Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(3 * base_board_length / 4, -base_board_length / 4, 0) '标注尺寸

            part.Extension.SelectByID2("D7@草图1@H472_2056.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
            Annotation = DisplayDimension.GetAnnotation()
            Annotation.SetPosition(2 * thick, -thick, 0) '移动标注
            part.ClearSelection2(True)
            part.Extension.SelectByID2("D8@草图1@H472_2056.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
            Annotation = DisplayDimension.GetAnnotation()
            Annotation.SetPosition(base_board_length - 3 * thick / 2, -thick, 0) '移动标注
            part.ClearSelection2(True)
            part.Extension.SelectByID2("D10@草图1@H472_2056.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            part.Extension.SelectByID2("D9@草图1@H472_2056.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
            part.Extension.SelectByID2("D1@草图1@H472_2056.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
            part.Extension.SelectByID2("D3@草图1@H472_2056.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
            part.Extension.SelectByID2("D6@草图1@H472_2056.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
            part.EditDelete()
            part.SketchManager.InsertSketch(True) '结束草图编辑
            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, width, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            P10 = part.SketchManager.CreatePoint(0, thick, 0)
            P11 = part.SketchManager.CreatePoint(0, -circle_edge_distance, 0)
            P10.Select4(False, Nothing)
            P11.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(-thick, -circle_edge_distance / 2, 0) '标注尺寸
            part.SketchManager.InsertSketch(True)

            part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            part.SketchManager.CreateCircleByRadius(-width / 2, -circle_edge_distance, 0, diameter / 2)
            DisplayDimension = part.AddDimension2(0, -circle_edge_distance, width / 2) '标注尺寸
            part.SketchManager.InsertSketch(True)
            part.FeatureCut(True, False, True, 2, 0, 0, 0, False, False, 0, 0, 0, 0, 0, 0) '反向切除,到下一面
        End Sub
        Public Sub Rear_hanging_climbing(length#, width#, height1#, height2#, height3#, thick#, band_fillet_radius#, base_circle_edge_distance#, base_center_distance#,
                        base_circle_diameter#, top_circle_edge_distance#, top_circle_diameter#, board_fillet_radius#, chamfer_length#)
            '后端吊攀(长,宽,高1,高2,高3,厚,折弯圆角半径,底圆边距,底圆中心距,底圆半径,顶圆边距,顶圆直径,板圆角半径,倒角长)

            part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            L1 = part.SketchManager.CreateLine(width - thick, 0, 0, width, 0, 0)
            L2 = part.SketchManager.CreateLine(width - thick, 0, 0, width - thick, height1, 0)
            L3 = part.SketchManager.CreateLine(width - thick, height1, 0, 0, height1 + height2, 0)
            L4 = part.SketchManager.CreateLine(0, height1 + height2, 0, 0, height1 + height2 + height3, 0)
            L5 = part.SketchManager.CreateLine(0, height1 + height2 + height3, 0, thick, height1 + height2 + height3, 0)

            L6 = part.SketchManager.CreateLine(width, 0, 0, width, height1, 0)
            L7 = part.SketchManager.CreateLine(thick, height1 + height2 + height3, 0, thick, height1 + height2, 0)
            L8 = part.SketchManager.CreateLine(thick, height1 + height2 - 0.001, 0, width / 2, height1 + height2 / 2, 0)
            L1Segment = L1
            L2Segment = L2
            L3Segment = L3
            L4Segment = L4
            L5Segment = L5

            L1Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(0, -thick, -(width + thick)) '标注尺寸
            L2Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(0, height1 / 2, thick) '标注尺寸
            L3Segment.Select4(False, Nothing)
            DisplayDimension = part.AddVerticalDimension2(0, height1 + height2 / 2, thick) '标注尺寸
            L3Segment.Select4(False, Nothing)
            DisplayDimension = part.AddHorizontalDimension2(0, -thick, -width / 2) '标注尺寸
            L4Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(0, height1 + height2 + height3 / 2, thick) '标注尺寸

            L1Segment.Select4(False, Nothing)
            L5Segment.Select4(True, Nothing)
            草图方法("平行")
            L1Segment.Select4(False, Nothing)
            L5Segment.Select4(True, Nothing)
            草图方法("相等")
            L6Segment = L6
            L7Segment = L7
            L8Segment = L8
            'L2Segment.Select4(False, Nothing)
            L7Segment.Select4(False, Nothing)
            草图方法("垂直")
            L6Segment.Select4(False, Nothing)
            草图方法("垂直")
            L3Segment.Select4(False, Nothing)
            L8Segment.Select4(True, Nothing)
            草图方法("平行")
            'L4Segment.Select4(False, Nothing)
            L3Segment.Select4(False, Nothing)
            L8Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(width / 2, height1 + height2 / 2, 0) '标注尺寸
            ChangeDimensionVaule(thick)
            'MsgBox(1)
            P2 = L2.GetEndPoint2
            P3 = L3.GetEndPoint2
            P6 = L6.GetEndPoint2
            P7 = L7.GetEndPoint2
            P8 = L8.GetStartPoint2
            P9 = L8.GetEndPoint2
            P7.Select4(False, Nothing)
            P8.Select4(True, Nothing)
            草图方法("合并")
            P6.Select4(False, Nothing)
            P9.Select4(True, Nothing)
            草图方法("合并")
            P2.Select4(False, Nothing)
            SketchManager.CreateFillet(band_fillet_radius, 1)
            P3.Select4(False, Nothing)
            SketchManager.CreateFillet(band_fillet_radius + thick, 1)
            P6.Select4(False, Nothing)
            SketchManager.CreateFillet(band_fillet_radius + thick, 1)
            P7.Select4(False, Nothing)
            SketchManager.CreateFillet(band_fillet_radius, 1)
            L1Segment.Select4(False, Nothing)
            L6Segment.Select4(True, Nothing)
            part.SketchManager.CreateChamfer(2, chamfer_length, 0) '倒角
            L4Segment.Select4(False, Nothing)
            L6Segment.Select4(True, Nothing)
            DisplayDimension = part.AddHorizontalDimension2(0, -2 * thick, -width / 2) '标注尺寸

            part.Extension.SelectByID2("D4@草图1@H472_2057.SLDPRT", "DIMENSION", 0.165327830771286, 0.0970706398837024, 0, False, 0, Nothing, 0)
            part.Extension.SelectByID2("D9@草图1@H472_2057.SLDPRT", "DIMENSION", 0.138455779065175, 0.140267692284992, 0, True, 0, Nothing, 1)
            part.Extension.SelectByID2("D6@草图1@H472_2057.SLDPRT", "DIMENSION", 0.180990219865215, 0.159212397024313, 0, True, 0, Nothing, 1)
            part.Extension.SelectByID2("D10@草图1@H472_2057.SLDPRT", "DIMENSION", 0.182225109117289, 0.178786429556451, 0, True, 0, Nothing, 1)
            part.Extension.SelectByID2("D8@草图1@H472_2057.SLDPRT", "DIMENSION", 0.197819075778652, 0.174075752127498, 0, True, 0, Nothing, 1)
            part.Extension.SelectByID2("D13@草图1@H472_2057.SLDPRT", "DIMENSION", 0.197819075778652, 0.174075752127498, 0, True, 0, Nothing, 1)
            part.EditDelete()
            part.Extension.SelectByID2("D13@草图1@H472_2057.SLDPRT", "DIMENSION", 0.197819075778652, 0.174075752127498, 0, False, 0, Nothing, 0)
            part.EditDelete()

            part.SketchManager.InsertSketch(True)
            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, length, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0) '拉伸,单向

            part.ClearSelection2(True)
            part.Extension.SelectByRay（-thick, 0, -(width - 3 * thick / 4), 1, 0, 0, 0.001, 1, False, 1, 0）
            part.Extension.SelectByRay（length + thick, 0, -(width - 3 * thick / 4), -1, 0, 0, 0.001, 1, True, 1, 0）
            part.Extension.SelectByRay（-thick, height1 + height2 + height3, -thick / 4, 1, 0, 0, 0.001, 1, False, 1, 0）
            part.Extension.SelectByRay（length + thick, height1 + height2 + height3, -thick / 4, -1, 0, 0, 0.001, 1, True, 1, 0）
            part.FeatureManager.FeatureFillet3(2, board_fillet_radius, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) '创建圆角特征
            part.Extension.SelectByID2("D1@圆角1@零件6.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
            Annotation = DisplayDimension.GetAnnotation()
            Annotation.SetPosition(-thick, board_fillet_radius, 0) '移动标注
            part.ClearSelection2(True)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            A1 = part.SketchManager.CreateCircleByRadius((length - base_center_distance) / 2, base_circle_edge_distance, 0, base_circle_diameter / 2) '用半径画圆
            A2 = part.SketchManager.CreateCircleByRadius((length + base_center_distance) / 2, base_circle_edge_distance, 0, base_circle_diameter / 2) '用半径画圆
            A3 = part.SketchManager.CreateCircleByRadius(length / 2, height1 + height2 + height3 - top_circle_edge_distance, 0, top_circle_diameter / 2) '用半径画圆
            A1Segment = A1
            A2Segment = A2
            A3Segment = A3
            A2Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(length + 2 * thick, height1, 0) '标注尺寸
            DisplayDimension.SetBrokenLeader2(False, 2) '直径标注类型,水平
            part.EditDimensionProperties2(0, 0, 0, "", "", True, 9, 2, True, 12, 12, "2x<MOD-DIAM>", "", True, "", "", False) '修改尺寸内容
            A3Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(length / 2, height1 + height2 + height3 - top_circle_edge_distance, 0) '标注尺寸
            A1Segment.Select4(False, Nothing)
            A2Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(length / 2, -thick, 0) '标注尺寸
            part.SketchManager.InsertSketch(True)
            part.FeatureCut(True, False, False, 1, 0, 0, 0, False, False, 0, 0, 0, 0, 0, 0) '正向贯穿切除,单向
            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            P1 = part.SketchManager.CreatePoint(0, 0, 0)
            P2 = part.SketchManager.CreatePoint(length, base_circle_edge_distance, 0)
            P3 = part.SketchManager.CreatePoint(length / 2, height1 + height2 + height3 - top_circle_edge_distance, 0)
            P4 = part.SketchManager.CreatePoint(length / 2, height1 + height2 + height3, 0)
            P1.Select4(False, Nothing)
            P2.Select4(True, Nothing)
            DisplayDimension = part.AddHorizontalDimension2(length / 2, -2 * thick, 0) '标注尺寸
            P1.Select4(False, Nothing)
            P2.Select4(True, Nothing)
            DisplayDimension = part.AddVerticalDimension2(length + thick, base_circle_edge_distance / 2, 0) '标注尺寸
            P3.Select4(False, Nothing)
            P4.Select4(True, Nothing)
            DisplayDimension = part.AddVerticalDimension2(length + thick, height1 + height2 + height3 - top_circle_edge_distance / 2, 0) '标注尺寸
            part.SketchManager.InsertSketch(True)
        End Sub
        Public Function ChangeDimensionVaule(vaule#) '修改标注尺寸
            Dimension = part.Parameter(DisplayDimension.GetNameForSelection())
            Dimension.SystemValue = vaule
        End Function
        Public Sub SetTolvalue(type#, max$, min$) '公差标注
            Dimension = DisplayDimension.GetDimension2(0)
            DimensionTolerance = Dimension.Tolerance
            DimensionTolerance.Type = type
            DimensionTolerance.SetFitValues(max, min)

        End Sub
        Public Function 草图方法(类型$)
            If 类型 = "固定" Then
                part.SketchAddConstraints("sgFIXED")
            ElseIf 类型 = "水平" Then
                part.SketchAddConstraints("sgHORIZONTAL2D")
            ElseIf 类型 = "点水平排布" Then
                part.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            ElseIf 类型 = "点竖直排布" Then
                part.SketchAddConstraints("sgVERTICALPOINTS2D")
            ElseIf 类型 = "垂直" Then
                part.SketchAddConstraints("sgVERTICAL2D")
            ElseIf 类型 = "共线" Then
                part.SketchAddConstraints("sgCOLINEAR")
            ElseIf 类型 = "全等" Then
                part.SketchAddConstraints("sgCORADIAL")
            ElseIf 类型 = "相互垂直" Then
                part.SketchAddConstraints("sgPERPENDICULAR")
            ElseIf 类型 = "平行" Then
                part.SketchAddConstraints("sgPARALLEL")
            ElseIf 类型 = "相切" Then
                part.SketchAddConstraints("sgTANGENT")
            ElseIf 类型 = "同心" Then
                part.SketchAddConstraints("sgCONCENTRIC")
            ElseIf 类型 = "重合" Then
                part.SketchAddConstraints("sgCOINCIDENT")
            ElseIf 类型 = "对称" Then
                part.SketchAddConstraints("sgSYMMETRIC")
            ElseIf 类型 = "相等" Then
                part.SketchAddConstraints("sgSAMELENGTH")
            ElseIf 类型 = "合并" Then
                part.SketchAddConstraints("sgMERGEPOINTS")
            ElseIf 类型 = "相等曲率" Then
                part.SketchAddConstraints("sgEQUALCURV3DALIGN")
            ElseIf 类型 = "曲线长度相等" Then
                part.SketchAddConstraints("sgSAMECURVELENGTH")
            ElseIf 类型 = "镜像" Then
                part.SketchMirror()
            End If

        End Function
        Public Function 基准轴(基准$) As SldWorks.Feature
            Dim AXIS As SldWorks.Feature
            Dim line1 As SldWorks.SketchLine
            Dim line1Segment As SldWorks.SketchSegment
            If 基准 = "Z" Then
                part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, -0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "Y" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "X" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0.01, 0, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            基准轴 = AXIS
        End Function
        Public Function Initial_setting(type#)
            If type = 0 Then '关闭捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 1 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 2 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 3 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            End If

        End Function
        Public Function 直径标注(类型$)
            If 类型 = "水平" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderHorizontalText) '直径水平
            ElseIf 类型 = "跟随" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swSolidLeaderAlignedText) '直径跟随标注线
            ElseIf 类型 = "断开" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderAlignedText) '标注与线断开
            End If
        End Function
        Public Sub [End](dz$)
            part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
            part.SaveAs3(dz, 0, 8)
        End Sub
        Public Sub set_attribute(name$, code$, material$)

            Dim cusproper As SldWorks.CustomPropertyManager
            cusproper = part.Extension.CustomPropertyManager("")
            cusproper.Set2("名称", name)
            cusproper.Set2("代号", code)
            cusproper.Set2("材料", material)
        End Sub
    End Class
    Public Class Supporting_board

        Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
        Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
        Dim NewDocument As SldWorks.ModelDoc2 = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_part.prtdot", 0, 0, 0)
        Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc '获取当前活动的文档对象
        Dim SketchManager As SldWorks.SketchManager = part.SketchManager '草图管理器
        Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager '声明变量获得特征管理器对象,用其接受特征管理器引用
        Dim Dimension As SldWorks.Dimension '模型的尺寸和公差
        Dim DisplayDimension As SldWorks.DisplayDimension '表示显示的尺寸实例
        Dim DimensionTolerance As DimensionTolerance
        Dim sketch As SldWorks.Sketch
        Dim Feature As SldWorks.Feature
        Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager '选择管理器的接口
        Const Pi As Double = 3.1415926535897931
        Public Sub Supporting_board(length#, right_height#, right_angle#, left_height#, left_angle#, thick#, House_shell_diameter#)
            '支板(长,右高,右角,左高,左角,厚,筒体直径)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            L1 = part.SketchManager.CreateLine(0, 0, 0, length, 0, 0)
            L2 = part.SketchManager.CreateLine(length, 0, 0, length - thick, right_height, 0)
            L3 = part.SketchManager.CreateLine(0, 0, 0, -thick, left_height, 0)
            L1Segment = L1
            L2Segment = L2
            L3Segment = L3
            L1Segment.Select4(False, Nothing)
            草图方法("水平")
            'DisplayDimension = part.AddDimension2(宽 / 2, -厚, 0) '标注尺寸
            'ChangeDimensionVaule(宽)
            P1 = L2.GetStartPoint2
            P2 = L2.GetEndPoint2
            P3 = L3.GetStartPoint2
            P4 = L3.GetEndPoint2
            P1.Select4(False, Nothing)
            P3.Select4(True, Nothing)
            草图方法("固定")
            L1Segment.Select4(False, Nothing)
            L2Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(length / 2, length / 2, 0)
            ChangeDimensionVaule(right_angle * Pi / 180)
            L2Segment.Select4(False, Nothing)
            DisplayDimension = part.AddVerticalDimension2(3 * length / 2, right_height / 2, 0)
            ChangeDimensionVaule(right_height)
            L1Segment.Select4(False, Nothing)
            L3Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(-2 * length / 3, length / 2, 0)
            ChangeDimensionVaule(left_angle * Pi / 180)
            L3Segment.Select4(False, Nothing)
            DisplayDimension = part.AddVerticalDimension2(-length / 2, left_height / 2, 0)
            ChangeDimensionVaule(left_height)
            A1Segment = part.SketchManager.Create3PointArc(P2.X, P2.Y, 0, P4.X, P4.Y, 0, length / 2, right_height / 2, 0)
            A1Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(0, right_height / 2, 0)
            ChangeDimensionVaule(House_shell_diameter / 2)
            part.Extension.SelectByID2("D5@草图1@S150_3350.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            part.EditDelete()
            part.SketchManager.InsertSketch(True) '结束草图编辑
            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, thick, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0)
            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            P1 = part.SketchManager.CreatePoint(0, 0, 0)
            P2 = part.SketchManager.CreatePoint(length, 0, 0)
            P1.Select4(False, Nothing)
            P2.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(length / 2, -thick, 0) '标注尺寸
            ChangeDimensionVaule(length)
            part.SketchManager.InsertSketch(True) '插入新的草图

        End Sub
        Public Function ChangeDimensionVaule(vaule#) '修改标注尺寸
            Dimension = part.Parameter(DisplayDimension.GetNameForSelection())
            Dimension.SystemValue = vaule
        End Function
        Public Sub SetTolvalue(type#, max$, min$) '公差标注
            Dimension = DisplayDimension.GetDimension2(0)
            DimensionTolerance = Dimension.Tolerance
            DimensionTolerance.Type = type
            DimensionTolerance.SetFitValues(max, min)

        End Sub
        Public Function 草图方法(类型$)
            If 类型 = "固定" Then
                part.SketchAddConstraints("sgFIXED")
            ElseIf 类型 = "水平" Then
                part.SketchAddConstraints("sgHORIZONTAL2D")
            ElseIf 类型 = "点水平排布" Then
                part.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            ElseIf 类型 = "点竖直排布" Then
                part.SketchAddConstraints("sgVERTICALPOINTS2D")
            ElseIf 类型 = "垂直" Then
                part.SketchAddConstraints("sgVERTICAL2D")
            ElseIf 类型 = "共线" Then
                part.SketchAddConstraints("sgCOLINEAR")
            ElseIf 类型 = "全等" Then
                part.SketchAddConstraints("sgCORADIAL")
            ElseIf 类型 = "相互垂直" Then
                part.SketchAddConstraints("sgPERPENDICULAR")
            ElseIf 类型 = "平行" Then
                part.SketchAddConstraints("sgPARALLEL")
            ElseIf 类型 = "相切" Then
                part.SketchAddConstraints("sgTANGENT")
            ElseIf 类型 = "同心" Then
                part.SketchAddConstraints("sgCONCENTRIC")
            ElseIf 类型 = "重合" Then
                part.SketchAddConstraints("sgCOINCIDENT")
            ElseIf 类型 = "对称" Then
                part.SketchAddConstraints("sgSYMMETRIC")
            ElseIf 类型 = "相等" Then
                part.SketchAddConstraints("sgSAMELENGTH")
            ElseIf 类型 = "合并" Then
                part.SketchAddConstraints("sgMERGEPOINTS")
            ElseIf 类型 = "相等曲率" Then
                part.SketchAddConstraints("sgEQUALCURV3DALIGN")
            ElseIf 类型 = "曲线长度相等" Then
                part.SketchAddConstraints("sgSAMECURVELENGTH")
            ElseIf 类型 = "镜像" Then
                part.SketchMirror()
            End If

        End Function
        Public Function 基准轴(基准$) As SldWorks.Feature
            Dim AXIS As SldWorks.Feature
            Dim line1 As SldWorks.SketchLine
            Dim line1Segment As SldWorks.SketchSegment
            If 基准 = "Z" Then
                part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, -0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "Y" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "X" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0.01, 0, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            基准轴 = AXIS
        End Function
        Public Function Initial_setting(type#)
            If type = 0 Then '关闭捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 1 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 2 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 3 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            End If

        End Function
        Public Function 直径标注(类型$)
            If 类型 = "水平" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderHorizontalText) '直径水平
            ElseIf 类型 = "跟随" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swSolidLeaderAlignedText) '直径跟随标注线
            ElseIf 类型 = "断开" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderAlignedText) '标注与线断开
            End If
        End Function
        Public Sub [End](dz$)
            part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
            part.SaveAs3(dz, 0, 8)
        End Sub
        Public Sub set_attribute(name$, code$, material$)

            Dim cusproper As SldWorks.CustomPropertyManager
            cusproper = part.Extension.CustomPropertyManager("")
            cusproper.Set2("名称", name)
            cusproper.Set2("代号", code)
            cusproper.Set2("材料", material)
        End Sub
    End Class
    Public Class Additive

        Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
        Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
        Dim NewDocument As SldWorks.ModelDoc2 = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_part.prtdot", 0, 0, 0)
        Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc '获取当前活动的文档对象
        Dim SketchManager As SldWorks.SketchManager = part.SketchManager '草图管理器
        Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager '声明变量获得特征管理器对象,用其接受特征管理器引用
        Dim Dimension As SldWorks.Dimension '模型的尺寸和公差
        Dim DisplayDimension As SldWorks.DisplayDimension '表示显示的尺寸实例
        Dim DimensionTolerance As DimensionTolerance
        Dim sketch As SldWorks.Sketch
        Dim Feature As SldWorks.Feature
        Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager '选择管理器的接口
        Const Pi As Double = 3.1415926535897931
        Public Sub Additive(length#, angle#, rop_circle_radius#, base_circle_diameter#, chamfer_length#, chamfer_angle#, thick#)
            '搭子(长,角度,顶圆半径,底圆直径,倒角长,倒角角度,厚)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            L01Segment = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
            L02Segment = part.SketchManager.CreateCenterLine(-0.01, 0, 0.01, 0, 0, 0)
            L1 = part.SketchManager.CreateLine(length / 4, length / 2, 0, length / 2, 0, 0)
            L2 = part.SketchManager.CreateLine(-length / 4, length / 2, 0, -length / 2, 0, 0)
            L1Segment = L1
            L2Segment = L2
            P1 = L1.GetStartPoint2
            P2 = L1.GetEndPoint2
            P3 = L2.GetStartPoint2
            P4 = L2.GetEndPoint2
            L02Segment.Select4(False, Nothing)
            草图方法("水平")
            P2.Select4(False, Nothing)
            L02Segment.Select4(True, Nothing)
            草图方法("重合")
            P4.Select4(False, Nothing)
            L02Segment.Select4(True, Nothing)
            草图方法("重合")
            L1Segment.Select4(False, Nothing)
            L01Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(length / 4, -length / 3, 0)
            ChangeDimensionVaule(angle * Pi / 360)
            P2.Select4(False, Nothing)
            L01Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(length / 4, 0, 0)
            ChangeDimensionVaule(length / 2)
            L1Segment.Select4(False, Nothing)
            L2Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(0, -length / 2, 0)
            ChangeDimensionVaule(angle * Pi / 180)
            P2.Select4(False, Nothing)
            P4.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(0, length, 0)
            ChangeDimensionVaule(length)
            A1 = part.SketchManager.Create3PointArc(P1.X, P1.Y, 0, P3.X, P3.Y, 0, 0, P3.Y + rop_circle_radius / 2, 0)
            A1Segment = A1
            A1Segment.Select4(False, Nothing)
            L2Segment.Select4(True, Nothing)
            草图方法("相切")
            A1Segment.Select4(False, Nothing)
            L1Segment.Select4(True, Nothing)
            草图方法("相切")
            A1Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(length / 4, length / 2, 0)
            ChangeDimensionVaule(rop_circle_radius)
            P5 = A1.GetStartPoint2
            P6 = A1.GetEndPoint2
            P1.Select4(False, Nothing)
            P5.Select4(True, Nothing)
            草图方法("合并")
            P3.Select4(False, Nothing)
            P6.Select4(True, Nothing)
            草图方法("合并")
            A2Segment = part.SketchManager.Create3PointArc(P2.X, P2.Y, 0, P4.X, P4.Y, 0, 0, 0.005, 0)
            A2Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(0, -length / 4, 0)
            ChangeDimensionVaule(base_circle_diameter)
            part.Extension.SelectByID2("D6@草图1@A472_2008.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            part.Extension.SelectByID2("D1@草图1@A472_2008.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
            part.Extension.SelectByID2("D2@草图1@A472_2008.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
            part.EditDelete()
            part.SketchManager.InsertSketch(True) '结束草图编辑
            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, thick, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0)
            part.Extension.SelectByRay（0, 0, 0, 0, 1, 0, 0.001, 1, False, 1, 0）
            part.Extension.SelectByRay（0, 0, thick, 0, 1, 0, 0.001, 1, True, 1, 0）
            part.FeatureManager.InsertFeatureChamfer(1, 1, chamfer_length, chamfer_angle * Pi / 180, 0, 0, 0, 0) '倒角

            part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            P7 = part.SketchManager.CreatePoint(0, 2 * chamfer_length, 0)
            P8 = part.SketchManager.CreatePoint(-thick, 2 * chamfer_length, 0)
            P7.Select4(False, Nothing)
            P8.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(0, length, thick / 2)
            part.SketchManager.InsertSketch(True)
            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            part.SketchManager.CreateCenterLine(0, -thick, 0, 0, length - 0.005, 0)
            part.SketchManager.InsertSketch(True)
            Dim Annotation As SldWorks.Annotation
            part.Extension.SelectByID2("D2@倒角1@A472_2008.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 1)
            DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
            Annotation = DisplayDimension.GetAnnotation()
            Annotation.SetPosition(0, chamfer_length / 2, -thick / 2) '移动标注
        End Sub

        Public Function ChangeDimensionVaule(vaule#) '修改标注尺寸
            Dimension = part.Parameter(DisplayDimension.GetNameForSelection())
            Dimension.SystemValue = vaule
        End Function
        Public Sub SetTolvalue(type#, max$, min$) '公差标注
            Dimension = DisplayDimension.GetDimension2(0)
            DimensionTolerance = Dimension.Tolerance
            DimensionTolerance.Type = type
            DimensionTolerance.SetFitValues(max, min)

        End Sub
        Public Function 草图方法(类型$)
            If 类型 = "固定" Then
                part.SketchAddConstraints("sgFIXED")
            ElseIf 类型 = "水平" Then
                part.SketchAddConstraints("sgHORIZONTAL2D")
            ElseIf 类型 = "点水平排布" Then
                part.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            ElseIf 类型 = "点竖直排布" Then
                part.SketchAddConstraints("sgVERTICALPOINTS2D")
            ElseIf 类型 = "垂直" Then
                part.SketchAddConstraints("sgVERTICAL2D")
            ElseIf 类型 = "共线" Then
                part.SketchAddConstraints("sgCOLINEAR")
            ElseIf 类型 = "全等" Then
                part.SketchAddConstraints("sgCORADIAL")
            ElseIf 类型 = "相互垂直" Then
                part.SketchAddConstraints("sgPERPENDICULAR")
            ElseIf 类型 = "平行" Then
                part.SketchAddConstraints("sgPARALLEL")
            ElseIf 类型 = "相切" Then
                part.SketchAddConstraints("sgTANGENT")
            ElseIf 类型 = "同心" Then
                part.SketchAddConstraints("sgCONCENTRIC")
            ElseIf 类型 = "重合" Then
                part.SketchAddConstraints("sgCOINCIDENT")
            ElseIf 类型 = "对称" Then
                part.SketchAddConstraints("sgSYMMETRIC")
            ElseIf 类型 = "相等" Then
                part.SketchAddConstraints("sgSAMELENGTH")
            ElseIf 类型 = "合并" Then
                part.SketchAddConstraints("sgMERGEPOINTS")
            ElseIf 类型 = "相等曲率" Then
                part.SketchAddConstraints("sgEQUALCURV3DALIGN")
            ElseIf 类型 = "曲线长度相等" Then
                part.SketchAddConstraints("sgSAMECURVELENGTH")
            ElseIf 类型 = "镜像" Then
                part.SketchMirror()
            End If

        End Function
        Public Function 基准轴(基准$) As SldWorks.Feature
            Dim AXIS As SldWorks.Feature
            Dim line1 As SldWorks.SketchLine
            Dim line1Segment As SldWorks.SketchSegment
            If 基准 = "Z" Then
                part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, -0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "Y" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "X" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0.01, 0, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            基准轴 = AXIS
        End Function
        Public Function Initial_setting(type#)
            If type = 0 Then '关闭捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 1 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 2 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 3 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            End If

        End Function
        Public Function 直径标注(类型$)
            If 类型 = "水平" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderHorizontalText) '直径水平
            ElseIf 类型 = "跟随" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swSolidLeaderAlignedText) '直径跟随标注线
            ElseIf 类型 = "断开" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderAlignedText) '标注与线断开
            End If
        End Function
        Public Sub [End](dz$)
            part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
            part.SaveAs3(dz, 0, 8)
        End Sub
        Public Sub set_attribute(name$, code$, material$)

            Dim cusproper As SldWorks.CustomPropertyManager
            cusproper = part.Extension.CustomPropertyManager("")
            cusproper.Set2("名称", name)
            cusproper.Set2("代号", code)
            cusproper.Set2("材料", material)
        End Sub
    End Class
    Public Class Cylinder

        Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
        Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
        Dim NewDocument As SldWorks.ModelDoc2 = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_part.prtdot", 0, 0, 0)
        Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc '获取当前活动的文档对象
        Dim SketchManager As SldWorks.SketchManager = part.SketchManager '草图管理器
        Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager '声明变量获得特征管理器对象,用其接受特征管理器引用
        Dim Dimension As SldWorks.Dimension '模型的尺寸和公差
        Dim DisplayDimension As SldWorks.DisplayDimension '表示显示的尺寸实例
        Dim DimensionTolerance As DimensionTolerance
        Dim sketch As SldWorks.Sketch
        Dim Feature As SldWorks.Feature
        Dim Face As Object()
        Dim swface As SldWorks.Face
        Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager '选择管理器的接口
        Dim swbodys As SldWorks.Body2
        Const Pi As Double = 3.1415926535897931
        Public Sub Cylinder(cylider_width#, outer_diameter#, thick#, hole_edge_distance#, hole_distance#, hole_length#, hole_width#, hole_fillet_radius#, houle_num#)
            '筒体(筒宽#, 外径#, 厚#, 孔边距#, 孔间距#, 孔长#, 孔宽#, 孔圆角#, 孔个数#)

            part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            L01Segment = part.SketchManager.CreateCenterLine(0, 0, 0, outer_diameter / 2 * Cos(89.5 * Pi / 180), outer_diameter / 2 * Sin(89.5 * Pi / 180), 0)
            L02Segment = part.SketchManager.CreateCenterLine(0, 0, 0, -outer_diameter / 2 * Cos(89.5 * Pi / 180), outer_diameter / 2 * Sin(89.5 * Pi / 180), 0)
            A1Segment = part.SketchManager.CreateCircleByRadius(0, 0, 0, outer_diameter / 2)
            A1Segment.Select4(False, Nothing)
            SketchManager.SketchTrim(0, 0, outer_diameter / 2, 0)

            A1Segment.Select4(False, Nothing)
            Dim customBendAllowanceData As Object '创建基体法兰
            part.FeatureManager.InsertSheetMetalBaseFlange2(thick, True, 0.001, cylider_width, 0.01, True, 0, 0, 1, customBendAllowanceData, False, 0, 0.0001, 0.0001, 0.5, True, False, True, True)
            part.Extension.SelectByRay(0, outer_diameter / 2, 0, 0, 0, 0.01, 0.001, 1, False, 1, 0) '展开基体法兰
            part.Extension.SelectByID2("基体折弯1", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            part.InsertSheetMetalUnfold()
            part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            L03Segment = part.SketchManager.CreateCenterLine(0, outer_diameter / 2, 0, -outer_diameter / 2, outer_diameter / 2, 0)
            L03Segment.Select4(False, Nothing)
            草图方法("水平")
            part.ClearSelection()

            Dim edge As Object
            Dim vEdgeCount As Integer
            part.Extension.SelectByID2("展开1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0) '选中特征,历遍所有面
            Feature = SelectionMgr.GetSelectedObject6(1, -1)
            Face = Feature.GetFaces()
            vEdgeCount = Feature.GetFaceCount()
            Debug.Print(vEdgeCount)
            part.ClearSelection2(True)

            'vEdgeCount = 0
            ''Do Until vEdgeCount > 100
            ''    For i = vEdgeCount To vEdgeCount + 5
            ''        selecmag.AddSelectionListObject(edge(i), Nothing)
            ''    Next
            ''    vEdgeCount = vEdgeCount + 5
            ''    MessageBox.Show(vEdgeCount)
            ''    part.ClearSelection2(True)
            ''    part.EditRebuild()
            ''Loop
            'Do Until vEdgeCount > 10
            '    SelectionMgr.AddSelectionListObject(Face(vEdgeCount), Nothing)

            '    MsgBox(vEdgeCount)
            '    part.ClearSelection2(True)
            '    vEdgeCount = vEdgeCount + 1
            '    part.EditRebuild()
            'Loop
            'part.ClearSelection2(True)

            SelectionMgr.AddSelectionListObject(Face(0), Nothing) '历遍面上所有边
            swface = SelectionMgr.GetSelectedObject6(1, -1)
            edge = swface.GetEdges()

            'vEdgeCount = swface.GetEdgeCount()
            'Debug.Print(vEdgeCount)
            'part.ClearSelection2(True)

            'vEdgeCount = 0
            ''Do Until vEdgeCount > 100
            ''    For i = vEdgeCount To vEdgeCount + 5
            ''        selecmag.AddSelectionListObject(edge(i), Nothing)
            ''    Next
            ''    vEdgeCount = vEdgeCount + 5
            ''    MessageBox.Show(vEdgeCount)
            ''    part.ClearSelection2(True)
            ''    part.EditRebuild()
            ''Loop
            'Do Until vEdgeCount > 20
            '    SelectionMgr.AddSelectionListObject(edge(vEdgeCount), Nothing)

            '    MsgBox(vEdgeCount)
            '    part.ClearSelection2(True)
            '    vEdgeCount = vEdgeCount + 1
            '    part.EditRebuild()
            'Loop
            part.ClearSelection2(True)
            SelectionMgr.AddSelectionListObject(edge(1), Nothing)
            SelectionMgr.AddSelectionListObject(edge(3), Nothing)
            L03Segment.Select4(True, Nothing)
            草图方法("对称")
            part.ClearSelection2(True)
            L03Segment.Select4(True, Nothing)
            草图方法("固定")

            L1 = part.SketchManager.CreateLine(-thick, thick, 0, -thick, 2 * thick, 0)
            L2 = part.SketchManager.CreateLine(-thick, 2 * thick, 0, -2 * thick, 2 * thick, 0)
            L3 = part.SketchManager.CreateLine(-2 * thick, 2 * thick, 0, -2 * thick, thick, 0)
            L4 = part.SketchManager.CreateLine(-2 * thick, thick, 0, -thick, thick, 0)
            L1Segment = L1
            L2Segment = L2
            L3Segment = L3
            L4Segment = L4
            P1 = L1.GetStartPoint2
            P2 = L1.GetEndPoint2
            P3 = L2.GetStartPoint2
            P4 = L2.GetEndPoint2
            P5 = L3.GetStartPoint2
            P6 = L3.GetEndPoint2
            P7 = L4.GetStartPoint2
            P8 = L4.GetEndPoint2
            P2.Select4(False, Nothing)
            P3.Select4(True, Nothing)
            草图方法("合并")

            L1Segment.Select4(False, Nothing)
            草图方法("垂直")
            L2Segment.Select4(False, Nothing)
            草图方法("水平")
            L1Segment.Select4(False, Nothing)
            L3Segment.Select4(True, Nothing)
            草图方法("平行")
            L1Segment.Select4(False, Nothing)
            L3Segment.Select4(True, Nothing)
            草图方法("相等")
            L2Segment.Select4(False, Nothing)
            L4Segment.Select4(True, Nothing)
            草图方法("平行")
            L2Segment.Select4(False, Nothing)
            L4Segment.Select4(True, Nothing)
            草图方法("相等")
            part.ClearSelection2(True)

            SelectionMgr.AddSelectionListObject(edge(2), Nothing)
            L1Segment.Select4(True, Nothing)
            DisplayDimension = part.AddHorizontalDimension2(hole_width, 0, -Pi * outer_diameter / 2 + 3 * hole_distance + hole_length)
            ChangeDimensionVaule(hole_edge_distance)
            L4Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(-hole_width / 2, 0, -Pi * outer_diameter / 2 + 3 * hole_distance + 2 * hole_length)
            ChangeDimensionVaule(hole_width)
            L3Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(-3 * hole_width / 2, 0, -Pi * outer_diameter / 2 + 2 * hole_distance + 3 * hole_length / 2)
            ChangeDimensionVaule(hole_length)
            L2Segment.Select4(False, Nothing)
            L03Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(hole_width, 0, -Pi * outer_diameter / 2 + hole_length)
            ChangeDimensionVaule(hole_length + 3 * hole_distance / 2)

            sketch = SketchManager.ActiveSketch '获得当前活动的草图对象
            SketchSegments = sketch.GetSketchSegments '获得此草图中的草图段
            For i = 0 To UBound(SketchSegments) '草图由很多个sketchsegment组成,for循环就是遍历草图中所有的sketchsegment
                SketchSegment = SketchSegments(i)
                [Boolean] = SketchSegment.Select4(False, Nothing) : Debug.Assert([Boolean])
                part.SketchConstraintsDelAll() '删除当前草图线段上的所有约束,仅删除当前所选草图线段,如果选择了两个或多个草图线段,则此方法不起作用
            Next i
            part.SketchConstraintsDelAll()
            P1.Select4(False, Nothing)
            P2.Select4(True, Nothing)
            P4.Select4(True, Nothing)
            P6.Select4(True, Nothing)
            SketchManager.CreateFillet(hole_fillet_radius, 1)
            L1Segment.SelectChain(False, Nothing) '选中线段链
            part.SketchManager.CreateLinearSketchStepAndRepeat(1, houle_num, 0, hole_distance + hole_length, 0, 90 * Pi / 180, "", False, False, False, False, False) '阵列草图段,不显示间距尺寸标注
            part.FeatureCut(True, False, True, 1, 0, 0, 0, False, False, 0, 0, 0, 0, 0, 0) '反向贯穿切除
            SelectionMgr.AddSelectionListObject(edge(1), Nothing)
            part.Extension.SelectByID2("基体折弯1", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            part.InsertSheetMetalFold() '折叠钣金件

            'part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            'part.SketchManager.InsertSketch(True) '插入新的草图
            'P1 = part.SketchManager.CreatePoint(0, 0, 0)
            'P2 = part.SketchManager.CreatePoint(-筒宽, 0, 0)
            'P1.Select4(False, Nothing)
            'P2.Select4(True, Nothing)
            'DisplayDimension = part.AddDimension2(-筒宽 / 2, -2 * 外径 / 3, 0)
            'part.SketchManager.InsertSketch(True)

            part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图
            L01Segment = part.SketchManager.CreateCenterLine(0, 0, 0, outer_diameter / 2 * Cos(89.5 * Pi / 180), outer_diameter / 2 * Sin(89.5 * Pi / 180), 0)
            L02Segment = part.SketchManager.CreateCenterLine(0, 0, 0, -outer_diameter / 2 * Cos(89.5 * Pi / 180), outer_diameter / 2 * Sin(89.5 * Pi / 180), 0)
            A1Segment = part.SketchManager.CreateCircleByRadius(0, 0, 0, outer_diameter / 2)
            DisplayDimension = part.AddDimension2(0, 2 * outer_diameter / 3 * Sin(60 * Pi / 180), 2 * outer_diameter / 3 * Cos(60 * Pi / 180))
            DisplayDimension.SetBrokenLeader2(False, 2) '直径标注类型,水平
            A1Segment.Select4(False, Nothing)
            SketchManager.SketchTrim(0, 0, outer_diameter / 2, 0)
            A2Segment = part.SketchManager.CreateCircleByRadius(0, 0, 0, outer_diameter / 2 - thick)
            DisplayDimension = part.AddDimension2(0, 2 * outer_diameter / 3 * Sin(30 * Pi / 180), 2 * outer_diameter / 3 * Cos(30 * Pi / 180))
            DisplayDimension.SetBrokenLeader2(False, 2) '直径标注类型,水平
            A2Segment.Select4(False, Nothing)
            SketchManager.SketchTrim(0, 0, outer_diameter / 2 - thick, 0)
            L01Segment.Select4(False, Nothing)
            L02Segment.Select4(True, Nothing)
            part.EditDelete()
            part.SketchManager.InsertSketch(True)
            part.ClearSelection2(True)


        End Sub


        Public Function 获得sketchsegment根据name(name$) As SldWorks.SketchSegment '画矩形时获得线段
            part.ClearSelection2(True)
            'line01.Select4(False, Nothing)
            sketch = SketchManager.ActiveSketch
            'points = sketch.GetSketchPoints2
            SketchSegments = sketch.GetSketchSegments
            For i = 0 To UBound(SketchSegments)
                SketchSegment = SketchSegments(i)
                'MsgBox(SketchSegment.name)
                '[Boolean] = SketchSegment.Select2(True, Nothing) : Debug.Assert([Boolean])
                If SketchSegment.GetName = name Then

                    获得sketchsegment根据name = SketchSegment
                End If
            Next i
            '获得sketchsegment = SketchSegment

            'Dim aaa As SldWorks.SketchSegment
            'aaa = 获得sketchsegment根据name("直线2")
            'aaa.Select2(False, 1)
            'MsgBox(99)
        End Function
        Public Function ChangeDimensionVaule(vaule#) '修改标注尺寸
            Dimension = part.Parameter(DisplayDimension.GetNameForSelection())
            Dimension.SystemValue = vaule
        End Function
        Public Sub SetTolvalue(type#, max$, min$) '公差标注
            Dimension = DisplayDimension.GetDimension2(0)
            DimensionTolerance = Dimension.Tolerance
            DimensionTolerance.Type = type
            DimensionTolerance.SetFitValues(max, min)

        End Sub
        Public Function 草图方法(类型$)
            If 类型 = "固定" Then
                part.SketchAddConstraints("sgFIXED")
            ElseIf 类型 = "水平" Then
                part.SketchAddConstraints("sgHORIZONTAL2D")
            ElseIf 类型 = "点水平排布" Then
                part.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            ElseIf 类型 = "点竖直排布" Then
                part.SketchAddConstraints("sgVERTICALPOINTS2D")
            ElseIf 类型 = "垂直" Then
                part.SketchAddConstraints("sgVERTICAL2D")
            ElseIf 类型 = "共线" Then
                part.SketchAddConstraints("sgCOLINEAR")
            ElseIf 类型 = "全等" Then
                part.SketchAddConstraints("sgCORADIAL")
            ElseIf 类型 = "相互垂直" Then
                part.SketchAddConstraints("sgPERPENDICULAR")
            ElseIf 类型 = "平行" Then
                part.SketchAddConstraints("sgPARALLEL")
            ElseIf 类型 = "相切" Then
                part.SketchAddConstraints("sgTANGENT")
            ElseIf 类型 = "同心" Then
                part.SketchAddConstraints("sgCONCENTRIC")
            ElseIf 类型 = "重合" Then
                part.SketchAddConstraints("sgCOINCIDENT")
            ElseIf 类型 = "对称" Then
                part.SketchAddConstraints("sgSYMMETRIC")
            ElseIf 类型 = "相等" Then
                part.SketchAddConstraints("sgSAMELENGTH")
            ElseIf 类型 = "合并" Then
                part.SketchAddConstraints("sgMERGEPOINTS")
            ElseIf 类型 = "相等曲率" Then
                part.SketchAddConstraints("sgEQUALCURV3DALIGN")
            ElseIf 类型 = "曲线长度相等" Then
                part.SketchAddConstraints("sgSAMECURVELENGTH")
            ElseIf 类型 = "镜像" Then
                part.SketchMirror()
            End If

        End Function
        Public Function 基准轴(基准$) As SldWorks.Feature
            Dim AXIS As SldWorks.Feature
            Dim line1 As SldWorks.SketchLine
            Dim line1Segment As SldWorks.SketchSegment
            If 基准 = "Z" Then
                part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, -0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "Y" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            If 基准 = "X" Then
                part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                part.SketchManager.InsertSketch(True)
                line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0.01, 0, 0)
                line1Segment = line1
                part.InsertSketch2(True)
                line1Segment.Select4(False, Nothing)
                part.InsertAxis2(True)
                AXIS = SelectionMgr.GetSelectedObject6(1, -1)
                AXIS.Select2(False, Nothing)
            End If
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            基准轴 = AXIS
        End Function
        Public Function Initial_setting(type#)
            If type = 0 Then '关闭捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 1 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 2 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 3 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            End If

        End Function
        Public Function 直径标注(类型$)
            If 类型 = "水平" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderHorizontalText) '直径水平
            ElseIf 类型 = "跟随" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swSolidLeaderAlignedText) '直径跟随标注线
            ElseIf 类型 = "断开" Then
                part.Extension.SetUserPreferenceString(swUserPreferenceStringValue_e.swDetailingLayer, swUserPreferenceOption_e.swDetailingDiameterDimension, "")
                part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimensionTextAndLeaderStyle, swUserPreferenceOption_e.swDetailingDiameterDimension, swDisplayDimensionLeaderText_e.swBrokenLeaderAlignedText) '标注与线断开
            End If
        End Function
        Public Sub [End](dz$)
            part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
            part.SaveAs3(dz, 0, 8)
        End Sub
        Public Sub set_attribute(name$, code$, material$)

            Dim cusproper As SldWorks.CustomPropertyManager
            cusproper = part.Extension.CustomPropertyManager("")
            cusproper.Set2("名称", name)
            cusproper.Set2("代号", code)
            cusproper.Set2("材料", material)
        End Sub
    End Class
    Public Class Cylinder_hole '筒体孔标注

        Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
        Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
        Dim NewDocument As SldWorks.ModelDoc2 = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_part.prtdot", 0, 0, 0)
        Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc '获取当前活动的文档对象
        Dim SketchManager As SldWorks.SketchManager = part.SketchManager '草图管理器
        Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager '声明变量获得特征管理器对象,用其接受特征管理器引用
        Dim Dimension As SldWorks.Dimension '模型的尺寸和公差
        Dim DisplayDimension As SldWorks.DisplayDimension '表示显示的尺寸实例
        Dim DimensionTolerance As DimensionTolerance
        Dim sketch As SldWorks.Sketch
        Dim Feature As SldWorks.Feature
        Dim Face As Object()
        Dim swface As SldWorks.Face
        Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager '选择管理器的接口
        Dim swbodys As SldWorks.Body2
        Const Pi As Double = 3.1415926535897931
        Public Sub Cylinder_hole(cylider_width#, thick#, hole_edge_distance#, hole_distance#, hole_length#, hole_width#, hole_fillet_radius#, houle_num#)
            '筒体孔标注(筒宽#,厚#, 孔边距#, 孔间距#, 孔长#, 孔宽#, 孔圆角#, 孔个数#)

            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中对象,在前视基准面
            part.SketchManager.InsertSketch(True) '插入新的草图

            L1 = part.SketchManager.CreateLine(0, 0, 0, 0, houle_num * hole_length + hole_distance * (houle_num + 3), 0)
            L2 = part.SketchManager.CreateLine(0, houle_num * (hole_length + hole_distance), 0, -cylider_width, houle_num * hole_length + hole_distance * (houle_num + 3), 0)
            L3 = part.SketchManager.CreateLine(-cylider_width, houle_num * (hole_length + hole_distance), 0, -cylider_width, 0, 0)
            L4 = part.SketchManager.CreateLine(-cylider_width, 0, 0, 0, 0, 0)
            L1Segment = L1
            L2Segment = L2
            L3Segment = L3
            L4Segment = L4

            L4Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(-hole_width, -hole_width, 0)
            L1Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(hole_width, hole_width, 0)

            L1Segment.Select4(False, Nothing)
            草图方法("垂直")
            L4Segment.Select4(False, Nothing)
            草图方法("水平")
            L1Segment.Select4(False, Nothing)
            草图方法("固定")
            L4Segment.Select4(False, Nothing)
            草图方法("固定")
            L1Segment.Select4(False, Nothing)
            L3Segment.Select4(True, Nothing)
            草图方法("平行")
            L1Segment.Select4(False, Nothing)
            L3Segment.Select4(True, Nothing)
            草图方法("相等")
            L2Segment.Select4(False, Nothing)
            L4Segment.Select4(True, Nothing)
            草图方法("平行")
            L2Segment.Select4(False, Nothing)
            L4Segment.Select4(True, Nothing)
            草图方法("相等")
            part.ClearSelection2(True)

            L5 = part.SketchManager.CreateLine(-thick, thick, 0, -thick, 2 * thick, 0)
            L6 = part.SketchManager.CreateLine(-thick, 2 * thick, 0, -2 * thick, 2 * thick, 0)
            L7 = part.SketchManager.CreateLine(-2 * thick, 2 * thick, 0, -2 * thick, thick, 0)
            L8 = part.SketchManager.CreateLine(-2 * thick, thick, 0, -thick, thick, 0)
            L5Segment = L5
            L6Segment = L6
            L7Segment = L7
            L8Segment = L8
            P1 = L5.GetStartPoint2
            P2 = L5.GetEndPoint2
            P3 = L6.GetStartPoint2
            P4 = L6.GetEndPoint2
            P5 = L7.GetStartPoint2
            P6 = L7.GetEndPoint2
            P7 = L8.GetStartPoint2
            P8 = L8.GetEndPoint2
            P2.Select4(False, Nothing)
            P3.Select4(True, Nothing)
            草图方法("合并")

            L5Segment.Select4(False, Nothing)
            草图方法("垂直")
            L6Segment.Select4(False, Nothing)
            草图方法("水平")
            L5Segment.Select4(False, Nothing)
            L7Segment.Select4(True, Nothing)
            草图方法("平行")
            L5Segment.Select4(False, Nothing)
            L7Segment.Select4(True, Nothing)
            草图方法("相等")
            L6Segment.Select4(False, Nothing)
            L8Segment.Select4(True, Nothing)
            草图方法("平行")
            L6Segment.Select4(False, Nothing)
            L8Segment.Select4(True, Nothing)
            草图方法("相等")
            part.ClearSelection2(True)

            L1Segment.Select4(False, Nothing)
            L5Segment.Select4(True, Nothing)
            DisplayDimension = part.AddHorizontalDimension2(hole_width, hole_width, 0)
            ChangeDimensionVaule(hole_edge_distance)
            L8Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(-hole_width, -hole_width, 0)
            ChangeDimensionVaule(hole_width)
            L7Segment.Select4(False, Nothing)
            DisplayDimension = part.AddDimension2(-3 * hole_width / 2, hole_length, 0)
            ChangeDimensionVaule(hole_length)
            L8Segment.Select4(False, Nothing)
            L4Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(hole_width, hole_length, 0)
            ChangeDimensionVaule(2 * hole_distance)

            sketch = SketchManager.ActiveSketch '获得当前活动的草图对象
            SketchSegments = sketch.GetSketchSegments '获得此草图中的草图段
            For i = 0 To UBound(SketchSegments) '草图由很多个sketchsegment组成,for循环就是遍历草图中所有的sketchsegment
                SketchSegment = SketchSegments(i)
                [Boolean] = SketchSegment.Select4(False, Nothing) : Debug.Assert([Boolean])
                part.SketchConstraintsDelAll() '删除当前草图线段上的所有约束,仅删除当前所选草图线段,如果选择了两个或多个草图线段,则此方法不起作用
            Next i
            part.SketchConstraintsDelAll()
            P1.Select4(False, Nothing)
            P2.Select4(True, Nothing)
            P4.Select4(True, Nothing)
            P6.Select4(True, Nothing)
            SketchManager.CreateFillet(hole_fillet_radius, 1)
            L5Segment.Select4(False, Nothing)
            L1Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(hole_width / 2, hole_length / 2 + 2 * hole_distance, 0)
            L5Segment.Select4(False, Nothing)
            L7Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(-hole_width / 2 - hole_edge_distance, hole_distance, 0)
            L6Segment.Select4(False, Nothing)
            L8Segment.Select4(True, Nothing)
            DisplayDimension = part.AddDimension2(-3 * hole_width / 2, hole_length / 2 + 2 * hole_distance, 0)

            Dim Annotation As Object
            part.Extension.SelectByID2("D1@草图1@零件1.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
            Annotation = DisplayDimension.GetAnnotation()
            Annotation.SetPosition(-hole_width / 2 - hole_edge_distance, hole_length / 3 + 2 * hole_distance, 0) '移动标注
            part.ClearSelection2(True)

            L5Segment.SelectChain(False, Nothing) '选中线段链
            part.SketchManager.CreateLinearSketchStepAndRepeat(1, houle_num, 0, hole_distance + hole_length, 0, 90 * Pi / 180, "", False, False, False, False, False) '阵列草图段,不显示间距尺寸标注

            If houle_num > 1 Then
                P1 = part.SketchManager.CreatePoint(-hole_width / 2 - hole_edge_distance, hole_length + 2 * hole_distance, 0)
                P2 = part.SketchManager.CreatePoint(-hole_width / 2 - hole_edge_distance, hole_length + 3 * hole_distance, 0)
                P1.Select4(False, Nothing)
                P2.Select4(True, Nothing)
                DisplayDimension = part.AddDimension2(hole_width / 2, hole_length + 5 * hole_distance / 2, 0)
            End If

            part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, 0.004, 0, False, False, 0, 0, 0, 0, 0, 0, 0, 0, True, False, True, 0, 0, 0) '拉伸,单向,
            part.ClearSelection2(True)
        End Sub
        Public Function ChangeDimensionVaule(vaule#) '修改标注尺寸
            Dimension = part.Parameter(DisplayDimension.GetNameForSelection())
            Dimension.SystemValue = vaule
        End Function
        Public Function 草图方法(类型$)
            If 类型 = "固定" Then
                part.SketchAddConstraints("sgFIXED")
            ElseIf 类型 = "水平" Then
                part.SketchAddConstraints("sgHORIZONTAL2D")
            ElseIf 类型 = "点水平排布" Then
                part.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            ElseIf 类型 = "点竖直排布" Then
                part.SketchAddConstraints("sgVERTICALPOINTS2D")
            ElseIf 类型 = "垂直" Then
                part.SketchAddConstraints("sgVERTICAL2D")
            ElseIf 类型 = "共线" Then
                part.SketchAddConstraints("sgCOLINEAR")
            ElseIf 类型 = "全等" Then
                part.SketchAddConstraints("sgCORADIAL")
            ElseIf 类型 = "相互垂直" Then
                part.SketchAddConstraints("sgPERPENDICULAR")
            ElseIf 类型 = "平行" Then
                part.SketchAddConstraints("sgPARALLEL")
            ElseIf 类型 = "相切" Then
                part.SketchAddConstraints("sgTANGENT")
            ElseIf 类型 = "同心" Then
                part.SketchAddConstraints("sgCONCENTRIC")
            ElseIf 类型 = "重合" Then
                part.SketchAddConstraints("sgCOINCIDENT")
            ElseIf 类型 = "对称" Then
                part.SketchAddConstraints("sgSYMMETRIC")
            ElseIf 类型 = "相等" Then
                part.SketchAddConstraints("sgSAMELENGTH")
            ElseIf 类型 = "合并" Then
                part.SketchAddConstraints("sgMERGEPOINTS")
            ElseIf 类型 = "相等曲率" Then
                part.SketchAddConstraints("sgEQUALCURV3DALIGN")
            ElseIf 类型 = "曲线长度相等" Then
                part.SketchAddConstraints("sgSAMECURVELENGTH")
            ElseIf 类型 = "镜像" Then
                part.SketchMirror()
            End If

        End Function
        Public Function Initial_setting(type#)
            If type = 0 Then '关闭捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 1 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 2 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            ElseIf type = 3 Then '激活捕捉,打开端点和草图点
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
                swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
            End If

        End Function
        Public Sub [End](dz$)
            part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
            part.SaveAs3(dz, 0, 8)
        End Sub


    End Class
End Class
Public Class Drawings
    Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
    Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
    Dim NewDocument As SldWorks.ModelDoc2
    Dim part As SldWorks.ModelDoc2 = swapp.ActiveDoc
    Dim SketchManager As SldWorks.SketchManager = part.SketchManager
    Dim FeatureManager As SldWorks.FeatureManager = part.FeatureManager
    Dim Dimension As SldWorks.Dimension
    Dim DisplayDimension As SldWorks.DisplayDimension
    Dim sketch As SldWorks.Sketch
    Dim Feature As SldWorks.Feature
    Dim SelectionMgr As SldWorks.SelectionMgr = part.SelectionManager
    Public Sub [End](dz$)
        part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3) '长度精度保留小数点后3位
        part.SaveAs3(dz, 0, 8)
    End Sub
    Public Sub Stator_pressing_ring_A4pDrw(dz$, Scale2#, front#, right#, up#, name$, material$, code$,
                                           frontX#, frontY#, rightX#, rightY#, upX#, upY#,
                                           skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$,
                                           outer_diameter#, resection_diameter#, roughness$, thinck_lable$)
        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)

        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager

        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见

        If front = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If
        If right = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*右视", 0, 0, 0)
        End If
        If up = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*上视", 0, 0, 0)
        End If


        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)
        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.164, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 52.5 / 1000 * 2 And Annpos(0) < 76.8 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText("江西兰叶科技有限公司")


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 12 / 1000 * 2 And Annpos(1) < 21.5 / 1000 * 2 Then
                        N1.SetText(name)


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 6.0 / 1000 * 2 And Annpos(1) < 12 / 1000 * 2 Then
                        N1.SetText(code)


                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)


                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                                 skills_requirement1 + Chr(13) + Chr(10) +
                                 skills_requirement2 + Chr(13) + Chr(10) +
                                 skills_requirement3 + Chr(13) + Chr(10) +
                                 skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.034, 0.09, 0)

        part.ActivateView("工程图视图1")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象
        Dim SF1, SF2 As Object '插入粗糙度
        SF1 = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", roughness, "", "RZ", "", "", "")
        SF1.GetAnnotation().SetPosition2(-outer_diameter / 2 * Cos(60 * PI / 180) / Scale2 + frontX, outer_diameter / 2 * Sin(60 * PI / 180) / Scale2 + frontY, 0)
        SF1.SetAngle(15 * PI / 180)
        SF2 = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", roughness, "", "RZ", "", "", "")
        SF2.GetAnnotation().SetPosition2(frontX, -resection_diameter / 2 / 4 + frontY, 0)
        SF2.SetAngle(0)

        Dim P1 As SketchPoint '插入标注
        P1 = part.SketchManager.CreatePoint(outer_diameter / 2 * Cos(60 * PI / 180 - 0.002), -outer_diameter / 2 * Sin(60 * PI / 180), 0)
        P1.Select4(False, Nothing)
        Dim myAnnotation As Object
        myNote = part.InsertNote(thinck_lable)
        myAnnotation = myNote.GetAnnotation()
        myAnnotation.SetLeader3(swLeaderStyle_e.swUNDERLINED, 0, True, False, False, False)
        myAnnotation.SetPosition(outer_diameter / 2 * Cos(60 * PI / 180) / Scale2 + frontX + 0.003, -outer_diameter / 2 * Sin(60 * PI / 180) / Scale2 + frontY - 0.005, 0)
        part.ClearSelection2(True)
    End Sub
    Public Sub Balance_ring_A4pDrw(dz$, Scale2#, front#, right#, up#, name$, material$, code$,
                                   frontX#, frontY#, rightX#, rightY#, upX#, upY#,
                                   skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$,
                                   outer_diameter#, resection_diameter#, thick#, roughness$)
        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)

        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager

        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见
        Draw = swapp.ActiveDoc

        If front = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If
        If right = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*右视", rightX#, rightY#, 0)
        End If
        If up = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*上视", upX#, upY#, 0)
        End If

        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)
        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.164, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 52.5 / 1000 * 2 And Annpos(0) < 76.8 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText("江西兰叶科技有限公司")


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 12 / 1000 * 2 And Annpos(1) < 21.5 / 1000 * 2 Then
                        N1.SetText(name)


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 6.0 / 1000 * 2 And Annpos(1) < 12 / 1000 * 2 Then
                        N1.SetText(code)


                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)


                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                                 skills_requirement1 + Chr(13) + Chr(10) +
                                 skills_requirement2 + Chr(13) + Chr(10) +
                                 skills_requirement3 + Chr(13) + Chr(10) +
                                 skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.034, 0.09, 0)

        part.ActivateView("工程图视图1")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象

        Dim SF1, SF2, SF3 As Object
        SF1 = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", "", "", "", roughness, "", "")
        SF1.GetAnnotation().SetPosition2(190 / 1000, 275 / 1000, 0)
        Dim L1s As SketchSegment
        L1s = part.SketchManager.CreateLine(0, -3 * outer_diameter / 4, 0, 0, 3 * outer_diameter / 4, 0)
        L1s.Select4(False, Nothing)
        part.CreateSectionViewAt5(rightX, rightY, 0, "B", 128, Nothing, 0) '创建剖面图

        part.ActivateView("工程图视图2")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象
        part.ActivateView("工程图视图2")
        part.Extension.SelectByID2("D1@草图3@工程图12-SectionAssembly-1-1@工程图视图2", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        part.Extension.SelectByID2("D1@草图8@工程图12-SectionAssembly-1-1@工程图视图2", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
        part.Extension.SelectByID2("D1@草图5@工程图12-SectionAssembly-1-1@工程图视图2", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
        part.Extension.SelectByID2("D1@草图4@工程图12-SectionAssembly-1-1@工程图视图2", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
        part.EditDelete()

        part.Extension.SelectByID2("D1@凸台-拉伸1@工程图13-SectionAssembly-1-1@工程图视图2", "DIMENSION", 0, 0, 0, False, 0, Nothing, 1)
        DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
        Annotation = DisplayDimension.GetAnnotation()
        Annotation.SetPosition(rightX + 6 * thick / Scale2, rightY - (outer_diameter / 2 + 5 * thick) / Scale2, 0) '移动标注

        Dim P1, P2 As SketchPoint
        P1 = part.SketchManager.CreatePoint(thick / 2, -（resection_diameter / 2 + （outer_diameter - resection_diameter） / 6）, 0)
        SF2 = part.Extension.InsertSurfaceFinishSymbol3(2, 0, 0, 0, 0, 0, 1, "", "", "", "", "", "", "") '添加粗糙度标识,不允许切削
        SF2.GetAnnotation().SetPosition2(P1.X, P1.Y, P1.Z)
        SF2.SetAngle(-90 * PI / 180)
        part.ActivateView("工程图视图2")
        P2 = part.SketchManager.CreatePoint(-thick / 2, -（resection_diameter / 2 + （outer_diameter - resection_diameter） / 6）, 0)
        SF3 = part.Extension.InsertSurfaceFinishSymbol3(2, 0, 0, 0, 0, 0, 1, "", "", "", "", "", "", "")
        SF3.GetAnnotation().SetPosition2(P2.X, P2.Y, P2.Z)
        SF3.SetAngle(90 * PI / 180)
    End Sub
    Public Sub Flange_A4pDrw(dz$, Scale2#, front#, right#, up#, name$, material$, code$,
                             frontX#, frontY#, rightX#, rightY#, upX#, upY#,
                             skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$,
                             outer_diameter#, roughness$)
        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)

        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager

        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见


        If front = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If
        If right = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*右视", rightX#, rightY#, 0)
        End If
        If up = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*上视", upX#, upY#, 0)
        End If


        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)
        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.164, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 52.5 / 1000 * 2 And Annpos(0) < 76.8 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText("江西兰叶科技有限公司")


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 12 / 1000 * 2 And Annpos(1) < 21.5 / 1000 * 2 Then
                        N1.SetText(name)


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 6.0 / 1000 * 2 And Annpos(1) < 12 / 1000 * 2 Then
                        N1.SetText(code)


                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)


                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                             skills_requirement1 + Chr(13) + Chr(10) +
                             skills_requirement2 + Chr(13) + Chr(10) +
                             skills_requirement3 + Chr(13) + Chr(10) +
                             skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.034, 0.09, 0)

        part.ActivateView("工程图视图1")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象
        Dim L1s As SketchSegment
        L1s = part.SketchManager.CreateLine(0, -3 * outer_diameter / 4, 0, 0, 3 * outer_diameter / 4, 0)
        L1s.Select4(False, Nothing)
        part.CreateSectionViewAt5(frontX, frontY, 0, "B", 132, Nothing, 0) '创建剖面图,并切换方向
        part.Extension.SelectByID2("工程图视图1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) '隐藏前视图
        part.SuppressView
        part.ClearSelection2(True)


        part.ActivateView("工程图视图2")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象
        part.Extension.SelectByID2("工程图视图2", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        part.ActivateView("工程图视图2")
        Dim myBreakLine As Object '创建断裂视图
        Dim myView1 As Object
        myView1 = part.SelectionManager.GetSelectedObject5(1)
        myBreakLine = myView1.InsertBreak3(1, -（outer_diameter / 2 - 0.07 * Scale2）, （outer_diameter / 2 - 0.07 * Scale2）, 5, 1, True) '"5"为锯齿状断裂线
        part.BreakView
        part.EditRebuild3()
        Dim SF As Object
        SF = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", "", "", "", roughness, "", "") '插入右上角粗糙度
        SF.GetAnnotation().SetPosition2(190 / 1000, 275 / 1000, 0)
        part.Extension.SelectByID2("D1@草图2@工程图3-SectionAssembly-1-1@工程图视图2", "DIMENSION", 0, 0, 0, False, 0, Nothing, 1)
        DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
        Annotation = DisplayDimension.GetAnnotation()
        Annotation.SetPosition(0.07, frontY, 0) '移动标注
        part.ClearSelection2(True)
        part.Extension.SelectByID2("D1@草图3@工程图3-SectionAssembly-1-1@工程图视图2", "DIMENSION", 0, 0, 0, False, 0, Nothing, 1)
        DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
        Annotation = DisplayDimension.GetAnnotation()
        Annotation.SetPosition(0.06, frontY, 0) '移动标注
        part.ClearSelection2(True)
        part.Extension.SelectByID2("D2@草图3@工程图3-SectionAssembly-1-1@工程图视图2", "DIMENSION", 0, 0, 0, False, 0, Nothing, 1)
        DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
        Annotation = DisplayDimension.GetAnnotation()
        Annotation.SetPosition(0.16, frontY, 0) '移动标注
        part.ClearSelection2(True)


    End Sub
    Public Sub Clsasp_A4pDrw(dz$, Scale2#, front#, right#, up#, name$, material$, code$,
                             frontX#, frontY#, rightX#, rightY#, upX#, upY#,
                             skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$,
                             length#)

        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)

        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager

        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见


        If front = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If
        If right = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*右视", rightX#, rightY#, 0)
        End If
        If up = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*上视", upX#, upY#, 0)
        End If


        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)
        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.164, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 52.5 / 1000 * 2 And Annpos(0) < 76.8 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText("江西兰叶科技有限公司")


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 12 / 1000 * 2 And Annpos(1) < 21.5 / 1000 * 2 Then
                        N1.SetText(name)


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 6.0 / 1000 * 2 And Annpos(1) < 12 / 1000 * 2 Then
                        N1.SetText(code)


                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)


                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                             skills_requirement1 + Chr(13) + Chr(10) +
                             skills_requirement2 + Chr(13) + Chr(10) +
                             skills_requirement3 + Chr(13) + Chr(10) +
                             skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.034, 0.09, 0)

        part.ActivateView("工程图视图1")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象

        part.ClearSelection2(True)
        part.Extension.SelectByID2("工程图视图1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        part.ActivateView("工程图视图1")
        Dim myBreakLine1 As Object
        Dim myView1 As Object
        myView1 = part.SelectionManager.GetSelectedObject5(1)
        myBreakLine1 = myView1.InsertBreak(0, -(length / 2 - 0.06), length / 2 - 0.06, 1) '断线方向,竖直为0
        part.BreakView
        part.Extension.SelectByID2("D2@草图1@C146_2046-1@工程图视图1", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        part.Extension.SelectByID2("D3@草图1@C146_2046-1@工程图视图1", "DIMENSION", 0, 0, 0, True, 0, Nothing, 1)
        part.EditDelete()
        part.ClearSelection2(True)
        Dim myBreakLine2 As Object
        Dim myView2 As Object
        part.Extension.SelectByID2("工程图视图2", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        part.ActivateView("工程图视图2")
        myView2 = part.SelectionManager.GetSelectedObject5(1)
        myBreakLine2 = myView2.InsertBreak(0, -(length / 2 - 0.06), length / 2 - 0.06, 1) '断线方向,竖直为0
        part.BreakView
        part.ActivateView("工程图视图2")
        part.Extension.SelectByID2("D1@凸台-拉伸1@C146_2046-2@工程图视图2", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
        Annotation = DisplayDimension.GetAnnotation()
        Annotation.SetPosition(0.2, upY, 0) '移动标注
    End Sub
    Public Sub Front_and_rear_panel_installation_platform_A4pDrw(dz$, Scale2#, front#, right#, up#, name$, material$, code$,
                                                                frontX#, frontY#, rightX#, rightY#, upX#, upY#,
                                                                 skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$,
                                                                 roughness$)
        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)

        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager

        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见


        If front = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If
        If right = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*右视", rightX#, rightY#, 0)
        End If
        If up = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*上视", upX#, upY#, 0)
        End If


        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)
        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.164, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 52.5 / 1000 * 2 And Annpos(0) < 76.8 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText("江西兰叶科技有限公司")


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 12 / 1000 * 2 And Annpos(1) < 21.5 / 1000 * 2 Then
                        N1.SetText(name)


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 6.0 / 1000 * 2 And Annpos(1) < 12 / 1000 * 2 Then
                        N1.SetText(code)


                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)


                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                                 skills_requirement1 + Chr(13) + Chr(10) +
                                 skills_requirement2 + Chr(13) + Chr(10) +
                                 skills_requirement3 + Chr(13) + Chr(10) +
                                 skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.034, 0.09, 0)

        part.ActivateView("工程图视图1")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象
        Dim SF As Object
        SF = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", "", "", "", roughness, "", "") '插入右上角粗糙度
        SF.GetAnnotation().SetPosition2(190 / 1000, 275 / 1000, 0)
        part.Extension.SelectByID2("D3@草图1@F126_2055-1@工程图视图1", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0) '删除尺寸标注
        part.EditDelete()


    End Sub
    Public Sub Outlet_board_installation_platform_A4pDrw(dz$, Scale2#, front#, right#, up#, name$, material$, code$,
                                                         frontX#, frontY#, rightX#, rightY#, upX#, upY#,
                                                         skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$,
                                                         thick#, height#, roughness$)
        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)

        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager

        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见


        If front = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If
        If right = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*左视", rightX#, rightY#, 0)
        End If
        If up = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*上视", upX#, upY#, 0)
        End If


        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)
        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.164, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 52.5 / 1000 * 2 And Annpos(0) < 76.8 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText("江西兰叶科技有限公司")


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 12 / 1000 * 2 And Annpos(1) < 21.5 / 1000 * 2 Then
                        N1.SetText(name)


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 6.0 / 1000 * 2 And Annpos(1) < 12 / 1000 * 2 Then
                        N1.SetText(code)


                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)


                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                                 skills_requirement1 + Chr(13) + Chr(10) +
                                 skills_requirement2 + Chr(13) + Chr(10) +
                                 skills_requirement3 + Chr(13) + Chr(10) +
                                 skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.034, 0.09, 0)

        part.ActivateView("工程图视图1")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象
        Dim SF, SF2, SF3 As Object
        SF = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", "", "", "", roughness, "", "") '插入右上角粗糙度
        SF.GetAnnotation().SetPosition2(190 / 1000, 275 / 1000, 0)
        part.ActivateView("工程图视图1")
        Dim P1, P2 As SketchPoint
        P1 = part.SketchManager.CreatePoint(0, height / 2, 0)
        SF2 = part.Extension.InsertSurfaceFinishSymbol3(2, 0, 0, 0, 0, 0, 1, "", "", "", "", "", "", "") '添加粗糙度标识,不允许切削
        SF2.GetAnnotation().SetPosition2(P1.X, P1.Y, P1.Z)
        SF2.SetAngle(0)
        part.ActivateView("工程图视图1")
        P2 = part.SketchManager.CreatePoint(0, height / 2 - thick, 0)
        SF3 = part.Extension.InsertSurfaceFinishSymbol3(2, 0, 0, 0, 0, 0, 1, "", "", "", "", "", "", "")
        SF3.GetAnnotation().SetPosition2(P2.X, P2.Y, P2.Z)
        SF3.SetAngle(180 * PI / 180)
        part.ActivateView("工程图视图2")
        part.Extension.SelectByID2("D2@草图1@O126_2061-2@工程图视图2", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        part.EditDelete()
    End Sub
    Public Sub Hanging_climbing_A4pDrw(dz$, Scale2#, front#, right#, up#, name$, material$, code$,
                                   frontX#, frontY#, leftX#, leftY#, upX#, upY#,
                                   skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$,
                                   circle_edge_distance#, roughness$)
        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)

        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager

        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见


        If front = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If
        If right = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*左视", leftX#, leftY#, 0)
        End If
        If up = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*上视", upX#, upY#, 0)
        End If


        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)
        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.164, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 52.5 / 1000 * 2 And Annpos(0) < 76.8 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText("江西兰叶科技有限公司")


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 12 / 1000 * 2 And Annpos(1) < 21.5 / 1000 * 2 Then
                        N1.SetText(name)


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 6.0 / 1000 * 2 And Annpos(1) < 12 / 1000 * 2 Then
                        N1.SetText(code)


                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)


                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                                 skills_requirement1 + Chr(13) + Chr(10) +
                                 skills_requirement2 + Chr(13) + Chr(10) +
                                 skills_requirement3 + Chr(13) + Chr(10) +
                                 skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.034, 0.09, 0)

        part.ActivateView("工程图视图1")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象
        Dim SF As Object
        SF = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", "", "", "", roughness, "", "") '插入右上角粗糙度
        SF.GetAnnotation().SetPosition2(190 / 1000, 275 / 1000, 0)
        part.ActivateView("工程图视图2")
        part.Extension.SelectByID2("D1@凸台-拉伸1@H472_2056-2@工程图视图2", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
        Annotation = DisplayDimension.GetAnnotation()
        Annotation.SetPosition(leftX#, leftY# + 3 * circle_edge_distance / Scale2 / 2, 0) '移动标注
        part.ClearSelection2(True)
        part.Extension.SelectByID2("D10@草图1@H472_2056-2@工程图视图2", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        part.EditDelete()
    End Sub
    Public Sub Rear_hanging_climbing_A4pDrw(dz$, Scale2#, front#, right#, up#, name$, material$, code$,
                                           frontX#, frontY#, leftX#, leftY#, upX#, upY#,
                                            skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$,
                                            width#, height1#, height2#, height3#, chamfer_length#, chamfer_label$, roughness$)
        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)

        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager

        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见


        If front = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If
        If right = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*左视", leftX#, leftY#, 0)
        End If
        If up = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*上视", upX#, upY#, 0)
        End If


        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)
        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.164, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 52.5 / 1000 * 2 And Annpos(0) < 76.8 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText("江西兰叶科技有限公司")


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 12 / 1000 * 2 And Annpos(1) < 21.5 / 1000 * 2 Then
                        N1.SetText(name)


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 6.0 / 1000 * 2 And Annpos(1) < 12 / 1000 * 2 Then
                        N1.SetText(code)


                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)


                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                                 skills_requirement1 + Chr(13) + Chr(10) +
                                 skills_requirement2 + Chr(13) + Chr(10) +
                                 skills_requirement3 + Chr(13) + Chr(10) +
                                 skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.034, 0.09, 0)

        part.ActivateView("工程图视图1")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象
        Dim SF As Object
        SF = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", "", "", "", roughness, "", "") '插入右上角粗糙度
        SF.GetAnnotation().SetPosition2(190 / 1000, 275 / 1000, 0)
        part.ActivateView("工程图视图1")
        part.Extension.SelectByID2("D13@草图1@H472_2057-1@工程图视图1", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        part.Extension.SelectByID2("D2@草图1@H472_2057-1@工程图视图1", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        part.Extension.SelectByID2("D1@凸台-拉伸1@H472_2057-1@工程图视图1", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        part.Extension.SelectByID2("D3@草图1@H472_2057-1@工程图视图1", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        part.Extension.SelectByID2("D5@草图1@H472_2057-1@工程图视图1", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        part.EditDelete()

        part.ActivateView("工程图视图2")
        Dim P1 As SldWorks.SketchPoint
        P1 = part.SketchManager.CreatePoint(-width / 2, -（height1 + height2 + height3) / 2 + chamfer_length, 0)
        P1.Select4(False, Nothing)
        'Dim myNote As Object
        Dim myAnnotation As Object
        myNote = part.InsertNote(chamfer_label)
        myNote.LockPosition = False
        myNote.Angle = 0
        myNote.SetBalloon(0, 0)
        myAnnotation = myNote.GetAnnotation()
        myAnnotation.SetLeader3(swLeaderStyle_e.swUNDERLINED, 0, True, False, False, False)
        myAnnotation.SetPosition(leftX - width / Scale2, leftY - （height2 + height3) / 2 / Scale2, 0)


    End Sub
    Public Sub Supporting_board_A4pDrw(dz$, Scale2#, front#, right#, up#, name$, material$, code$,
                                       frontX#, frontY#, rightX#, rightY#, upX#, upY#,
                                       skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$,
                                       thick#, roughness$)
        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)

        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager

        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见


        If front = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If
        If right = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*右视", rightX#, rightY#, 0)
        End If
        If up = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*上视", upX#, upY#, 0)
        End If


        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)
        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.164, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 52.5 / 1000 * 2 And Annpos(0) < 76.8 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText("江西兰叶科技有限公司")


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 12 / 1000 * 2 And Annpos(1) < 21.5 / 1000 * 2 Then
                        N1.SetText(name)


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 6.0 / 1000 * 2 And Annpos(1) < 12 / 1000 * 2 Then
                        N1.SetText(code)


                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)


                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                                 skills_requirement1 + Chr(13) + Chr(10) +
                                 skills_requirement2 + Chr(13) + Chr(10) +
                                 skills_requirement3 + Chr(13) + Chr(10) +
                                 skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.034, 0.09, 0)

        part.ActivateView("工程图视图1")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象
        Dim SF, SF2, SF3 As Object
        SF = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", "", "", "", roughness, "", "") '插入右上角粗糙度
        SF.GetAnnotation().SetPosition2(190 / 1000, 275 / 1000, 0)
        part.ActivateView("工程图视图2")
        Dim P1, P2 As SketchPoint
        P1 = part.SketchManager.CreatePoint(0, thick / 2, 0)
        SF2 = part.Extension.InsertSurfaceFinishSymbol3(2, 0, 0, 0, 0, 0, 1, "", "", "", "", "", "", "") '添加粗糙度标识,不允许切削
        SF2.GetAnnotation().SetPosition2(P1.X, P1.Y, P1.Z)
        SF2.SetAngle(0)
        part.ActivateView("工程图视图2")
        P2 = part.SketchManager.CreatePoint(0, -thick / 2, 0)
        SF3 = part.Extension.InsertSurfaceFinishSymbol3(2, 0, 0, 0, 0, 0, 1, "", "", "", "", "", "", "")
        SF3.GetAnnotation().SetPosition2(P2.X, P2.Y, P2.Z)
        SF3.SetAngle(180 * PI / 180)
    End Sub
    Public Sub Additive_A4pDrw(dz$, Scale2#, front#, right#, up#, name$, material$, code$,
                               frontX#, frontY#, rightX#, rightY#, upX#, upY#,
                               skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$, roughness$)
        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)

        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager

        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见


        If front = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If
        If right = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*右视", rightX#, rightY#, 0)
        End If
        If up = 1 Then
            myView = Draw.CreateDrawViewFromModelView3(dz, "*上视", upX#, upY#, 0)
        End If


        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)
        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.164, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 52.5 / 1000 * 2 And Annpos(0) < 76.8 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 21.5 / 1000 * 2 And Annpos(1) < 30.5 / 1000 * 2 Then
                        N1.SetText("江西兰叶科技有限公司")


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 12 / 1000 * 2 And Annpos(1) < 21.5 / 1000 * 2 Then
                        N1.SetText(name)


                    End If
                    If Annpos(0) > 76.8 / 1000 * 2 And Annpos(0) < 102.5 / 1000 * 2 And Annpos(1) > 6.0 / 1000 * 2 And Annpos(1) < 12 / 1000 * 2 Then
                        N1.SetText(code)


                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)


                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                                 skills_requirement1 + Chr(13) + Chr(10) +
                                 skills_requirement2 + Chr(13) + Chr(10) +
                                 skills_requirement3 + Chr(13) + Chr(10) +
                                 skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.034, 0.09, 0)

        part.ActivateView("工程图视图1")
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64 + 128, True, False, True, True) '导入要显示的对象
        Dim SF As Object
        SF = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", "", "", "", roughness, "", "") '插入右上角粗糙度
        SF.GetAnnotation().SetPosition2(190 / 1000, 275 / 1000, 0)
        part.Extension.SelectByID2("D1@凸台-拉伸1@A472_2008-2@工程图视图2", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        part.EditDelete()


    End Sub
    Public Sub Cylinder_A3Drw(dz$, Scale2#, front#, right#, unfold#, name$, material$, code$,
                               frontX#, frontY#, rightX#, rightY#, frontXh#, frontYh#,
                              skills_requirement_title$, skills_requirement1$, skills_requirement2$, skills_requirement3$, skills_requirement4$,
                              cylider_width#, hole_width#, hole_length#, houle_num#, outer_diameter#, thick#, dzh$)

        Dim Draw As SldWorks.DrawingDoc
        Dim View As SldWorks.View '视图对象
        Dim Notes As Object
        Dim Count As Long
        Dim Annpos() As Double
        Dim Annotation As SldWorks.Annotation
        Dim myView As SldWorks.View
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a3.drwdot", 1, 0, 0)


        Dim DrawTitle As String
        DrawTitle = Draw.GetTitle
        Debug.Print(DrawTitle)
        'Dim myViewname As String
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager
        part.Extension.SetUserPreferenceInteger(372, 208, 2) '直径半径折弯标注
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 4)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, swDisplayMode_e.swHIDDEN_GREYED)
        swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swHiddenEdgeDisplayDefault, 1) '隐藏线可见
        Dim myView1, myview2, myview3 As SldWorks.View
        If front = 1 Then
            myView1 = Draw.CreateDrawViewFromModelView3(dz, "*前视", frontX#, frontY#, 0)
        End If

        If right = 1 Then
            myview2 = Draw.CreateDrawViewFromModelView3(dz, "*右视", rightX#, rightY#, 0)
        End If

        If unfold = 1 Then
            myview3 = Draw.CreateDrawViewFromModelView3(dzh, "*前视", frontXh#, frontYh#, 0)
        End If

        part.ClearSelection2(True)
        'myViewname = myView.Name
        'Debug.Print(myViewname)
        'myView = Draw.CreateDrawViewFromModelView3(dz1 + dz2, "*右视", 0.182, 0.192, 0)
        'myViewname = myView.Name
        'Debug.Print(
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowLength, 0, 0.004)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0, 0.002)
        part.Extension.SetUserPreferenceDouble(swUserPreferenceDoubleValue_e.swDetailingArrowHeight, 0, 0.0005)

        Dim myNote As Object
        myNote = part.InsertNote("公司名称")
        myNote.GetAnnotation().SetPosition(0.375, 0.055, 0)
        part.ClearSelection2(True)

        View = Draw.GetFirstView
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 315 / 1000 And Annpos(0) < 365 / 1000 And Annpos(1) > 43 / 1000 And Annpos(1) < 61 / 1000 Then
                        N1.SetText(material)
                    End If


                    If Annpos(0) > 365 / 1000 And Annpos(0) < 415 / 1000 And Annpos(1) > 42 / 1000 And Annpos(1) < 61 / 1000 Then
                        N1.SetText("江西兰叶科技有限公司")
                    End If


                    If Annpos(0) > 365 / 1000 And Annpos(0) < 415 / 1000 And Annpos(1) > 23 / 1000 And Annpos(1) < 43 / 1000 Then
                        N1.SetText(name)
                    End If

                    If Annpos(0) > 365 / 1000 And Annpos(0) < 415 / 1000 And Annpos(1) > 14 / 1000 And Annpos(1) < 23 / 1000 Then
                        N1.SetText(code)
                    End If


                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(code)
                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop

        Dim Sheet1 As SldWorks.Sheet '图纸对象
        Dim SheetPr() As Double
        Sheet1 = Draw.GetCurrentSheet()
        SheetPr = Sheet1.GetProperties2()
        SheetPr(2) = 1
        SheetPr(3) = Scale2
        SheetPr(4) = False
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild3()
        myNote = part.InsertNote(skills_requirement_title + Chr(13) + Chr(10) +
                                 skills_requirement1 + Chr(13) + Chr(10) +
                                 skills_requirement2 + Chr(13) + Chr(10) +
                                 skills_requirement3 + Chr(13) + Chr(10) +
                                 skills_requirement4）
        myNote.GetAnnotation().SetPosition(0.26, 0.09, 0)


        Dim SF, SF2, SF3 As Object
        SF = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", "", "", "", "12.5", "", "") '插入右上角粗糙度
        SF.GetAnnotation().SetPosition2(0.39, 0.275, 0)

        part.ActivateView("工程图视图1")
        part.Extension.SelectByID2(myView1.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        part.InsertModelAnnotations3(0, 32768, True, False, True, True)
        part.EditUndo2(1)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64, True, False, True, True)

        part.ActivateView("工程图视图2")
        part.Extension.SelectByID2(myview2.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64, False, False, True, True)

        part.ActivateView("工程图视图1")
        Dim P1, P2 As SketchPoint
        P1 = part.SketchManager.CreatePoint(hole_width, outer_diameter / 2, 0)
        SF2 = part.Extension.InsertSurfaceFinishSymbol3(2, 0, 0, 0, 0, 0, 1, "", "", "", "", "", "", "") '添加粗糙度标识,不允许切削
        SF2.GetAnnotation().SetPosition2(P1.X, P1.Y, P1.Z)
        SF2.SetAngle(0)
        part.ActivateView("工程图视图1")
        P2 = part.SketchManager.CreatePoint(-hole_width, outer_diameter / 2 - thick, 0)
        SF3 = part.Extension.InsertSurfaceFinishSymbol3(2, 0, 0, 0, 0, 0, 1, "", "", "", "", "", "", "")
        SF3.GetAnnotation().SetPosition2(P2.X, P2.Y, P2.Z)
        SF3.SetAngle(180 * PI / 180)

        part.ActivateView("工程图视图1")
        part.Extension.SelectByID2("D2@基体-法兰1@H004_2163-1@工程图视图1", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        DisplayDimension = SelectionMgr.GetSelectedObject6(1, -1)
        Annotation = DisplayDimension.GetAnnotation()
        Annotation.SetPosition(frontX, frontY - outer_diameter / 2 / Scale2 - 0.005, 0) '移动标注

        Dim P3 As SketchPoint '插入注解
        P3 = part.SketchManager.CreatePoint(cylider_width / 3, -outer_diameter / 2, 0)
        P3.Select4(False, Nothing)
        Dim myAnnotation As Object
        myNote = part.InsertNote("k")
        myNote.LockPosition = False
        myNote.Angle = 0
        myNote.SetBalloon(0, 0)
        myAnnotation = myNote.GetAnnotation()
        myAnnotation.SetLeader3(swLeaderStyle_e.swSTRAIGHT, 0, True, False, False, False)
        myAnnotation.SetPosition(frontX + cylider_width / 3 / Scale2, frontY - outer_diameter / 2 / Scale2 - 0.005, 0)

        part.ActivateView("工程图视图3") '创建局部视图
        part.SketchManager.CreateCircleByRadius(cylider_width / 2 - hole_width, 0, 0, hole_length * (houle_num / 2 + 1))
        part.CreateDetailViewAt4(frontXh, frontYh, 0, 0, 1, 2 * Scale2, "K向展开", 0, False, False, False, False)
        part.Extension.SelectByID2("工程图视图3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) '隐藏前视图
        part.SuppressView
        part.Extension.SelectByID2("工程图视图4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) '显示局部视图
        part.UnsuppressView

        part.ActivateView("工程图视图4")
        part.InsertModelAnnotations3(0, 32768 + 32 + 4 + 2 + 64, False, False, True, True)
        part.Extension.SelectByID2("细节项目260@工程图视图4", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
        Notes = SelectionMgr.GetSelectedObject6(1, -1)
        Annotation = Notes.GetAnnotation()
        Annotation.SetPosition(frontXh - cylider_width / 2 / Scale2 / 2, frontYh, 0) '移动标注
        part.ClearSelection2(True)
    End Sub


End Class

Public Class Assembly
    Dim swapp As SldWorks.SldWorks = CreateObject("Sldworks.application")
    Dim OpenDoc7 As SldWorks.ModelDoc2 = swapp.OpenDoc7("C:\Users\Public\Desktop\SOLIDWORKS 2019.lnk")
    Dim NewDocument As SldWorks.ModelDoc2
    Dim part As SldWorks.ModelDoc2
    Dim SketchManager As SldWorks.SketchManager
    Dim FeatureManager As SldWorks.FeatureManager
    Dim SelectionMgr As SldWorks.SelectionMgr

    Dim AssemblyDoc As SldWorks.AssemblyDoc
    Dim AssemblyTitle As String
    Dim Component2 As SldWorks.Component2
    Dim facenumber, facenumber1, facenumber2, facenumber3, facenumber4, facenumber5, facenumber6, facenumber7, facenumber8, facenumber9, facenumber10, facenumber11, facenumber12 As Integer
    Dim edgenumber, edgenumber1, edgenumber2, edgenumber3, edgenumber4, edgenumber5, edgenumber6, edgenumber7, edgenumber8, edgenumber9, edgenumber10, edgenumber11, edgenumber12 As Integer
    Dim errors As Long
    Dim longwarnings As Long
    Dim swbody, swbody1, swbody2, swbody3, swbody4, swbody5, swbody6, swbody7, swbody8, swbody9, swbody10 As SldWorks.Body2

    Dim swface As SldWorks.Face2

    Dim Dimension As SldWorks.Dimension
    Dim DisplayDimension As SldWorks.DisplayDimension
    Dim sketch As SldWorks.Sketch
    Dim Feature, Feature1, Feature2, Feature3, sketchfeature1, sketchfeature2, sketchfeature3, 基准轴feature1, 基准轴feature2, 基准轴feature3 As SldWorks.Feature
    Dim DFeature, DFeature1, DFeature2, DFeature3, Dsketchfeature1, Dsketchfeature2, Dsketchfeature3 As SldWorks.Feature

    Dim Annotation As SldWorks.Annotation
    Dim Annotations() As Object
    Dim DimensionTolerance As SldWorks.DimensionTolerance

    Dim Draw As SldWorks.DrawingDoc
    Dim View As SldWorks.View '视图对象
    Dim Notes As Object
    Dim Count As Long
    Dim Annpos() As Double
    Dim mySFSymbol As SldWorks.SFSymbol

    Dim myView As SldWorks.View
    Dim myNote As Object
    Dim face As SldWorks.Face2
    Dim Object数组1(), Object数组2(), Object数组3(), face1(), face2(), face3(), face4(), face5(), face6(), face7(), face8(), face9(), edge1(), edge2(), edge3(), edge4(), edge5(), edge6(), edge7(), edge8(), edge9() As Object
    Dim Object1, Object2, Object3 As Object

    Public A1, A2, A3, A4, A5, A6, A7, A8, A9, A10 As SldWorks.SketchArc
    Public L1, L2, L3, L4, L5, L6, L7, L8, L9, L10, L11, L12, L13, L14, L15, L01, L02, L03, L04, L05 As SldWorks.SketchLine
    Public DL1, DL2, DL3, DL4, DL5, DL6 As SldWorks.SketchLine
    Public A1Segment, A2Segment, A3Segment, A4Segment, A5Segment, L1Segment, L2Segment, L3Segment, L4Segment, L5Segment, L6Segment, L7Segment,
            L8Segment, L9Segment, L10Segment, L11Segment, L12Segment, L13Segment, L14Segment, L15Segment, L01Segment, L02Segment, L03Segment, L04Segment, L05Segment, DL1Segment, DL2Segment,
            DL3Segment, DL4Segment, DL5Segment, DL6Segment, SketchSegment, SketchSegment1, SketchSegment2, SketchSegment3 As SldWorks.SketchSegment
    Public SketchSegments(), points() As Object
    Public [Boolean] As Boolean
    Public P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, DP1, DP2, DP3, DP4, DP5, DP6, DP7, DP8, DP9, DP10, DP11, DP12, DP13, DP14, DP15 As SldWorks.SketchPoint
    Public X1, X2, X3, Y1, Y2, Y3, Z1, Z2, Z3 As Double
    Public Note As SldWorks.Note

    Dim RenderMaterial As SldWorks.RenderMaterial
    Dim swConfig As SldWorks.Configuration
    Dim status As Boolean
    Dim Entity As SldWorks.Entity






    Dim mysqlcon As MySqlConnection
    Dim mysqlcom As MySqlCommand
    Dim dr As MySqlDataReader

    Public Sub Assembly_initialization() '装配图初始化
        swapp.Visible = True
        swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_assembly.asmdot", 0, 0, 0)
        AssemblyDoc = swapp.ActiveDoc
        AssemblyTitle = AssemblyDoc.GetTitle
        part = swapp.ActiveDoc
        SketchManager = part.SketchManager
        FeatureManager = part.FeatureManager
        SelectionMgr = part.SelectionManager
        part.Extension.InsertScene("\scenes\01 basic scenes\00 soft box.p2s")

        swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swEdgesDefaultBulkSelection, True)
        swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swEdgesHiddenEdgeSelectionInWireframe, True)

        swapp.SetUserPreferenceIntegerValue(SwConst.swUserPreferenceIntegerValue_e.swEdgesTangentEdgeDisplay, 3) '切边不可见
        'swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swDisplayGraphicsComponents, True)
        part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True) '隐藏草图
        swapp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUseFolderAsDefaultSearchLocation, False) '设置异形孔向导/toolbox取消打勾
        'swapp.GetM

    End Sub

    Public Sub Rotor_core_assembly() '转子铁芯装配


        Dim mysqlcon As MySqlConnection
        Dim mysqlcom As MySqlCommand
        Dim dr As MySqlDataReader
        mysqlcon = New MySqlConnection("server=localhost" & ";userid=root" & ";password=123456" & ";database=eq_214_2431;pooling=false")
        '//打开数据库连接
        mysqlcon.Open()
        '//sql查询
        mysqlcom = New MySqlCommand("select * from _5eq_676_2105", mysqlcon)
        dr = mysqlcom.ExecuteReader()
        dr.Read()
        Do Until dr.GetString("id") = 1
            dr.Read()
        Loop



        Dim Code As String = dr("Code")
        Dim name As String = dr("name")
        Dim company As String = dr("company")
        Dim material As String = dr("material")

        Dim _1 As SldWorks.Component2
        Dim _2 As SldWorks.Component2
        Dim _3 As SldWorks.Component2
        Dim _4 As SldWorks.Component2


        For i = 1 To 4
            swapp.OpenDoc6(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)

            Component2 = part.AddComponent5(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0 + (i - 1) * 0.4)
            swapp.CloseDoc(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"))

            Select Case i > 0
                Case i = 1
                    _1 = Component2
                Case i = 2
                    _2 = Component2
                Case i = 3
                    _3 = Component2
                Case i = 4
                    _4 = Component2
            End Select
        Next


        For i = 1 To 4
            Select Case i > 0
                Case i = 1
                    Component2 = _1
                Case i = 2
                    Component2 = _2
                Case i = 3
                    Component2 = _3
                Case i = 4
                    Component2 = _4
            End Select
            part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
            swbody = SelectionMgr.GetSelectedObject6(1, -1)
            Select Case i > 0
                Case i = 1
                    swbody1 = swbody
                Case i = 2
                    swbody2 = swbody
                Case i = 3
                    swbody3 = swbody
                Case i = 4
                    swbody4 = swbody
            End Select
            part.ClearSelection2(True)
        Next





        part.ClearSelection2(True) '隐藏所有草图
        For i = 1 To 4
            Select Case i > 0
                Case i = 1
                    Component2 = _1
                Case i = 2
                    Component2 = _2
                Case i = 3
                    Component2 = _3
                Case i = 4
                    Component2 = _4
            End Select
            For j = 1 To 30
                part.Extension.SelectByID2("草图" + j.ToString + "@" + Component2.Name2() + "@" + AssemblyTitle, "SKETCH", 0, 0, 0, True, 0, Nothing, 0)
            Next
        Next
        'MsgBox(999)
        part.BlankSketch()


        'MsgBox(666)
        '

        渲染（_1, "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\data\graphics\materials\Metal\Steel\polished steel.p2m"） '剖光钢
        渲染（_2, "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\data\graphics\materials\Metal\Copper\matte copper.p2m"） '无光红铜
        渲染（_3, "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\data\graphics\materials\Metal\Copper\matte copper.p2m"） '无光红铜
        渲染（_4, "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\data\graphics\materials\Metal\Steel\polished steel.p2m"） '剖光钢


        part.UnblankSketch()


        face1 = swbody1.GetFaces
        face2 = swbody2.GetFaces
        face3 = swbody3.GetFaces
        face4 = swbody4.GetFaces

        edge1 = swbody1.GetEdges
        edge2 = swbody2.GetEdges
        edge3 = swbody3.GetEdges
        edge4 = swbody4.GetEdges


        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(face1(0), Nothing)
        SelectionMgr.AddSelectionListObject(face2(1), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(face1(91), Nothing)
        SelectionMgr.AddSelectionListObject(face3(0), Nothing)
        配合（"同轴", 1）
        SelectionMgr.AddSelectionListObject(face3(2), Nothing)
        SelectionMgr.AddSelectionListObject(face2(0), Nothing)
        配合（"重合", 0）
        SelectionMgr.AddSelectionListObject(face1(39), Nothing)
        SelectionMgr.AddSelectionListObject(face4(0), Nothing)
        配合（"同轴", 1）
        SelectionMgr.AddSelectionListObject(face2(0), Nothing)
        SelectionMgr.AddSelectionListObject(face4(2), Nothing)
        配合（"距离", 0, 0.07, False, 1, 0）


        _2.Select2(False, 1)
        part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, True, 2, Nothing, 0) '镜像零件
        Dim compsToMirror(0) As SldWorks.Component2
        Dim swMirrorPlane As SldWorks.Feature
        compsToMirror(0) = SelectionMgr.GetSelectedObjectsComponent4(1, 1) '加入标记
        swMirrorPlane = SelectionMgr.GetSelectedObject6(1, 2)
        AssemblyDoc.MirrorComponents3(swMirrorPlane, Nothing, 1, True, compsToMirror, True, "", 2, Nothing, "", 1, False, True, False)


        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(edge2（111）, Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        part.Extension.SelectByID2(_4.Name2 & "@*", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
        part.FeatureManager.FeatureCircularPattern5(4, 2 * PI, False, "NULL", False, True, False, False, False, False, 1, 0, "NULL", False)

        '遍历（1, swbody2.GetEdges, swbody2.GetFaceCount）

        Dim _3_1 As SldWorks.Feature
        Dim _3_2 As SldWorks.Feature
        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(edge2（26）, Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        part.Extension.SelectByID2(_3.Name2 & "@*", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
        _3_1 = part.FeatureManager.FeatureCircularPattern5(3, -7.8 * PI / 180, False, "NULL", False, False, False, False, False, False, 1, 0, "NULL", False)

        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(edge2（26）, Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        part.Extension.SelectByID2(_3.Name2 & "@*", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
        _3_1.Select2(True, 1)
        'part.Extension.SelectByID2("局部圆周阵列2", "COMPPATTERN", 0, 0, 0, True, 1, Nothing, 0)
        _3_2 = part.FeatureManager.FeatureCircularPattern5(2, -4 * 7.8 * PI / 180, False, "NULL", False, False, False, False, False, False, 1, 0, "NULL", False) '圆周阵列零件,角度,数量

        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(edge2（111）, Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        part.Extension.SelectByID2(_3.Name2 & "@*", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
        _3_1.Select2(True, 1)
        _3_2.Select2(True, 1)
        part.FeatureManager.FeatureCircularPattern5(4, 2 * PI, False, "NULL", False, True, False, False, False, False, 1, 0, "NULL", False) '圆周阵列零件,等间距

        part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True) '隐藏草图


        '遍历（1, swbody4.GetFaces, swbody4.GetFaceCount）


        Dim cusproper As SldWorks.CustomPropertyManager
        cusproper = part.Extension.CustomPropertyManager("")
        cusproper.Set2("名称", "转子铁芯")
        cusproper.Set2("代号", dr("code"))
        cusproper.Set2("材料", " ")
        part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True) '隐藏标注

        part.SaveAs3(dr("dz1") + dr("code") + dr("dz4"), 0, 8)







        '工程图

        A3模板(" ", dr("company"), dr("name"), dr("Code"))
        myView = Draw.CreateDrawViewFromModelView3(dr("dz1") + dr("code") + dr("dz4"), "*前视", 0.38, 0.14, 0)
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager
        part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitSystem, 0, SwConst.swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(372, 204, 2)
        part.Extension.SetUserPreferenceInteger(516, 2, 0)
        part.Extension.SetUserPreferenceInteger(517, 208, 0)
        part.Extension.SetUserPreferenceInteger(372, 208, 2) '直径半径折弯标注
        设置图纸比例(dr("proportion"))
        part.ViewDisplayHidden '隐藏线不可见

        part.ActivateView("工程图视图1") '创建剪裁视图
        part.SketchManager.CreateCornerRectangle(0, 0, 0, -0.18, 0.18, 0)
        'Dim myView As Object
        myView = part.ActiveDrawingView
        myView.Crop()


        Dim 分母 As Double
        Dim a As Object
        a = 索引字符串(dr("proportion"), ":")
        分母 = a(1)

        Dim myView1 As View
        part.ActivateView("工程图视图1")
        Dim L1s As SketchSegment
        L1s = part.SketchManager.CreateLine(0, 0, 0, 0, 0.08 * 分母, 0)
        L1s.Select4(False, Nothing)
        myView1 = part.CreateSectionViewAt5(0.16, 0, 0, " ", 128, Nothing, 0) '创建剖面图,并切换方向




        '添加尺寸
        part.ActivateView("工程图视图2")
        P1 = part.SketchManager.CreatePoint(-dr("width_l") - dr("thick") - dr("left"), dr("hight2") - dr("hight1") / 2, 0)
        P2 = part.SketchManager.CreatePoint(-dr("width_l") - dr("thick"), dr("hight1") / 2, 0)
        P3 = part.SketchManager.CreatePoint(-dr("width_l"), dr("hight1") / 2, 0)
        P4 = part.SketchManager.CreatePoint(dr("width_r"), dr("hight1") / 2, 0)
        P5 = part.SketchManager.CreatePoint(dr("width_r") + dr("thick"), dr("hight1") / 2, 0)
        P6 = part.SketchManager.CreatePoint(dr("width_r") + dr("thick") + dr("right"), dr("hight2") - dr("hight1") / 2, 0)

        P1.Select4(False, Nothing)
        P2.Select4(True, Nothing)
        DisplayDimension = part.AddHorizontalDimension2(0.16 - (dr("width_l") - dr("thick") - dr("left") / 2) / 分母, 0.1745 + dr("hight1") / 2 / 分母, 0)
        P3.Select4(False, Nothing)
        P2.Select4(True, Nothing)
        DisplayDimension = part.AddHorizontalDimension2(0.16 - (dr("width_l") - dr("thick") - dr("left") / 2) / 分母, 0.1745 + 2 * dr("hight1") / 3 / 分母, 0)
        P3.Select4(False, Nothing)
        P4.Select4(True, Nothing)
        DisplayDimension = part.AddHorizontalDimension2(0.16, 0.1745 + 2 * dr("hight1") / 3 / 分母, 0)
        P5.Select4(False, Nothing)
        P4.Select4(True, Nothing)
        DisplayDimension = part.AddHorizontalDimension2(0.16 - (dr("width_r") + dr("thick") + dr("right") / 2) / 分母, 0.1745 + 2 * dr("hight1") / 3 / 分母, 0)
        P5.Select4(False, Nothing)
        P6.Select4(True, Nothing)
        DisplayDimension = part.AddHorizontalDimension2(0.16 - (dr("width_r") + dr("thick") + dr("right") / 2) / 分母, 0.1745 + dr("hight1") / 2 / 分母, 0)
        part.ClearSelection2(True)



        Object数组1 = 索引字符串（dr（"skills_requirement"））
        技术要求（0.075, 0.085, Object数组1（0）, Object数组1（1）, Object数组1（2）, Object数组1（3）, Object数组1（4））
        part.ClearSelection2(True)

        插入零件标号(myView1)

        插入BOM表(myView, 1)


        part.SaveAs3(dr("dz1") + dr("code") + dr("dz3"), 0, 8)

        dr.Close()
        mysqlcom.Dispose()
        mysqlcon.Close()
        mysqlcon.Dispose()

    End Sub
    Public Sub Rotor_core_with_winding_assembly() '带绕组转子铁芯装配


        Dim mysqlcon As MySqlConnection
        Dim mysqlcom As MySqlCommand
        Dim dr As MySqlDataReader
        mysqlcon = New MySqlConnection("server=localhost" & ";userid=root" & ";password=123456" & ";database=eq_214_2431;pooling=false")
        '//打开数据库连接
        mysqlcon.Open()
        '//sql查询
        mysqlcom = New MySqlCommand("select * from _5eq_526_2137", mysqlcon)
        dr = mysqlcom.ExecuteReader()
        dr.Read()
        Do Until dr.GetString("id") = 1
            dr.Read()
        Loop



        Dim Code As String = dr("Code")
        Dim name As String = dr("name")
        Dim company As String = dr("company")
        Dim material As String = dr("material")

        Dim _5EQ_676_2105 As SldWorks.Component2
        Dim _5EQ_123_2000 As SldWorks.Component2
        Dim _8EQ_769_2231 As SldWorks.Component2
        Dim _8EQ_769_2232 As SldWorks.Component2
        Dim _8EQ_168_2007_1 As SldWorks.Component2
        Dim _8EQ_168_2007_2 As SldWorks.Component2


        '插入装配体
        For i = 1 To 2
            swapp.OpenDoc6(dr("dz1") + dr("Asm" + i.ToString + " Code") + dr("dz4"), 2, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("dz1") + dr("Asm" + i.ToString + " Code") + dr("dz4"), 0, "", False, "", 0, 0, 0 + (i - 1) * 0.4)
            swapp.CloseDoc(dr("dz1") + dr("Asm" + i.ToString + " Code") + dr("dz4"))

            Select Case i > 0
                Case i = 1
                    _5EQ_676_2105 = Component2
                Case i = 2
                    _5EQ_123_2000 = Component2

            End Select
        Next






        '插入零件
        For i = 3 To 4
            swapp.OpenDoc6(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0 + (i + 1) * 0.4)
            swapp.CloseDoc(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"))

            Select Case i > 0
                Case i = 1
                    _8EQ_769_2231 = Component2
                Case i = 2
                    _8EQ_769_2232 = Component2
                Case i = 3
                    _8EQ_168_2007_1 = Component2
                Case i = 4
                    _8EQ_168_2007_2 = Component2
            End Select
        Next





        For i = 3 To 4
            Select Case i > 0
                Case i = 1
                    Component2 = _8EQ_769_2231
                Case i = 2
                    Component2 = _8EQ_769_2232
                Case i = 3
                    Component2 = _8EQ_168_2007_1
                Case i = 4
                    Component2 = _8EQ_168_2007_2

            End Select
            part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
            swbody = SelectionMgr.GetSelectedObject6(1, -1)
            Select Case i > 0
                Case i = 1
                    swbody1 = swbody
                Case i = 2
                    swbody2 = swbody
                Case i = 3
                    swbody3 = swbody
                Case i = 4
                    swbody4 = swbody

            End Select
            part.ClearSelection2(True)
        Next

        'face1 = swbody1.GetFaces
        'face2 = swbody2.GetFaces
        face3 = swbody3.GetFaces
        face4 = swbody4.GetFaces

        'edge1 = swbody1.GetEdges
        'edge2 = swbody2.GetEdges
        edge3 = swbody3.GetEdges
        edge4 = swbody4.GetEdges



        Dim body1(), body2(), body3() As Object

        part.Extension.SelectByID2(dr("Asm1 Code") + "-1@*/" + dr("A1P1 Code") + "-1@" + dr("Asm1 Code"), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Component2 = SelectionMgr.GetSelectedObject6(1, -1)
        body1 = Component2.GetBodies3(0, 0)
        face5 = body1(0).GetFaces
        edge5 = body1(0).GetEdges


        part.Extension.SelectByID2(dr("Asm2 Code") + "-1@*/" + dr("A2P1 Code") + "-1@" + dr("Asm2 Code"), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Component2 = SelectionMgr.GetSelectedObject6(1, -1)
        body2 = Component2.GetBodies3(0, 0)
        face6 = body2(0).GetFaces
        edge6 = body2(0).GetEdges

        part.Extension.SelectByID2(dr("Asm2 Code") + "-1@*/" + dr("A2P2 Code") + "-1@" + dr("Asm2 Code"), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Component2 = SelectionMgr.GetSelectedObject6(1, -1)
        body3 = Component2.GetBodies3(0, 0)
        face7 = body3(0).GetFaces
        edge7 = body3(0).GetEdges


        '遍历（1, edge6, 100, 0）
        SelectionMgr.AddSelectionListObject(edge7(0), Nothing)
        SelectionMgr.AddSelectionListObject(face5(142), Nothing)
        配合（"重合", 0）
        SelectionMgr.AddSelectionListObject(edge6(0), Nothing)
        SelectionMgr.AddSelectionListObject(face5(143), Nothing)
        配合（"重合", 0）

        SelectionMgr.AddSelectionListObject(face3(4), Nothing)
        SelectionMgr.AddSelectionListObject(face6(4), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(edge6(0), Nothing)
        SelectionMgr.AddSelectionListObject(edge3(5), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(edge6(11), Nothing)
        SelectionMgr.AddSelectionListObject(edge3(4), Nothing)
        配合（"重合", 0） '0是同向对齐
        SelectionMgr.AddSelectionListObject(face4(4), Nothing)
        SelectionMgr.AddSelectionListObject(face7(4), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(edge7(0), Nothing)
        SelectionMgr.AddSelectionListObject(edge4(7), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(edge7(19), Nothing)
        SelectionMgr.AddSelectionListObject(edge4(4), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(face4(0), Nothing)
        SelectionMgr.AddSelectionListObject(face5(142), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(face5(0), Nothing)
        SelectionMgr.AddSelectionListObject(face4(3), Nothing)
        配合（"距离", 0, 0.1, True）


        '遍历（1, face4, 400, 0）

        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(edge5（111）, Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        'part.Extension.SelectByID2(_3.Name2 & "@*", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
        _5EQ_123_2000.Select2(True, 1)
        _8EQ_168_2007_1.Select2(True, 1)
        _8EQ_168_2007_2.Select2(True, 1)
        part.FeatureManager.FeatureCircularPattern5(4, 2 * PI, False, "NULL", False, True, False, False, False, False, 1, 0, "NULL", False) '圆周阵列零件,等间距










        Dim cusproper As SldWorks.CustomPropertyManager
        cusproper = part.Extension.CustomPropertyManager("")
        cusproper.Set2("名称", "带绕组转子铁芯")
        cusproper.Set2("代号", dr("code"))
        cusproper.Set2("材料", " ")
        part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True) '隐藏标注
        part.SaveAs3(dr("dz1") + dr("code") + dr("dz4"), 0, 8)

        '工程图

        A3模板(" ", dr("company"), dr("name"), dr("Code"))
        myView = Draw.CreateDrawViewFromModelView3(dr("dz1") + dr("code") + dr("dz4"), "*前视", 0.38, 0.14, 0)
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager
        part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitSystem, 0, SwConst.swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(372, 204, 2)
        part.Extension.SetUserPreferenceInteger(516, 2, 0)
        part.Extension.SetUserPreferenceInteger(517, 208, 0)
        part.Extension.SetUserPreferenceInteger(372, 208, 2) '直径半径折弯标注
        设置图纸比例(dr("proportion"))
        part.ViewDisplayHidden '隐藏线不可见

        Dim 分母 As Double
        Dim a As Object
        a = 索引字符串(dr("proportion"), ":")
        分母 = a(1)

        part.ActivateView("工程图视图1") '创建剪裁视图
        part.SketchManager.CreateCornerRectangle(0, 0, 0, -0.08 * 分母, 0.08 * 分母, 0)
        myView = part.ActiveDrawingView
        myView.Crop()

        Dim myView1 As View
        part.ActivateView("工程图视图1")
        Dim L1s As SketchSegment
        L1s = part.SketchManager.CreateLine(0, 0, 0, 0, 0.08 * 分母, 0)
        L1s.Select4(False, Nothing)
        myView1 = part.CreateSectionViewAt5(0.16, 0, 0, " ", 128, Nothing, 0) '创建剖面图,并切换方向




        插入零件标号(myView)

        插入零件标号(myView1)


        插入BOM表(myView1, 2)

        part.SaveAs3(dr("dz1") + dr("code") + dr("dz3"), 0, 8)


        dr.Close()
        mysqlcom.Dispose()
        mysqlcon.Close()
        mysqlcon.Dispose()
    End Sub
    Public Sub Exciter_rotor_core() '励磁机转子铁芯
        Dim mysqlcon As MySqlConnection
        Dim mysqlcom As MySqlCommand
        Dim dr As MySqlDataReader
        mysqlcon = New MySqlConnection("server=localhost" & ";userid=root" & ";password=123456" & ";database=eq_214_2431;pooling=false")
        '//打开数据库连接
        mysqlcon.Open()
        '//sql查询
        mysqlcom = New MySqlCommand("select * from _5eq_682_2018", mysqlcon)
        dr = mysqlcom.ExecuteReader()
        dr.Read()
        Do Until dr.GetString("id") = 1
            dr.Read()
        Loop



        Dim Code As String = dr("Code")
        Dim name As String = dr("name")
        Dim company As String = dr("company")
        Dim material As String = dr("material")


        Dim _1 As SldWorks.Component2
        Dim _2 As SldWorks.Component2
        Dim _3 As SldWorks.Component2
        Dim _4 As SldWorks.Component2

        '插入零件
        For i = 1 To 1
            swapp.OpenDoc6(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0 + (i - 1) * 0.4)
            swapp.CloseDoc(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"))

            Select Case i > 0
                Case i = 1
                    _1 = Component2
                Case i = 2
                    _2 = Component2
                Case i = 3
                    _3 = Component2
                Case i = 4
                    _4 = Component2
            End Select
        Next
        For i = 3 To 4
            swapp.OpenDoc6(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0 + (i - 1) * 0.4)
            swapp.CloseDoc(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"))

            Select Case i > 0
                Case i = 1
                    _1 = Component2
                Case i = 2
                    _2 = Component2
                Case i = 3
                    _3 = Component2
                Case i = 4
                    _4 = Component2
            End Select
        Next


        For i = 1 To 3
            Select Case i > 0
                Case i = 1
                    Component2 = _1
                Case i = 2
                    Component2 = _3
                Case i = 3
                    Component2 = _4


            End Select
            part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
            swbody = SelectionMgr.GetSelectedObject6(1, -1)
            Select Case i > 0
                Case i = 1
                    swbody1 = swbody
                Case i = 2
                    swbody3 = swbody
                Case i = 3
                    swbody4 = swbody

            End Select
            part.ClearSelection2(True)
        Next
        face1 = swbody1.GetFaces
        ' face2 = swbody2.GetFaces
        face3 = swbody3.GetFaces
        face4 = swbody4.GetFaces

        edge1 = swbody1.GetEdges
        'edge2 = swbody2.GetEdges
        edge3 = swbody3.GetEdges
        edge4 = swbody4.GetEdges







        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(edge1(65), Nothing)
        SelectionMgr.AddSelectionListObject(edge4(20), Nothing)
        配合（"重合", 0）
        SelectionMgr.AddSelectionListObject(edge1(54), Nothing)
        SelectionMgr.AddSelectionListObject(edge3(0), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(edge1(158), Nothing)
        SelectionMgr.AddSelectionListObject(edge3(5), Nothing)
        配合（"重合", 0）
        SelectionMgr.AddSelectionListObject(face4(8), Nothing)
        SelectionMgr.AddSelectionListObject(face3(1), Nothing)
        配合（"重合", 1）
        ''MsgBox(1)
        '遍历（1, swbody3.GetFaces, 400）

        SelectionMgr.AddSelectionListObject(edge1(54), Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        _4.Select2(True, 1)
        part.FeatureManager.FeatureCircularPattern5(6, 2 * PI, False, "NULL", False, True, False, False, False, False, 1, 0, "NULL", False) '圆周阵列零件,等间距




        Dim cusproper As SldWorks.CustomPropertyManager
        cusproper = part.Extension.CustomPropertyManager("")
        cusproper.Set2("名称", "励磁机转子铁芯")
        cusproper.Set2("代号", dr("code"))
        cusproper.Set2("材料", " ")

        part.SaveAs3(dr("dz1") + dr("code") + dr("dz4"), 0, 8)






        '工程图
        A4P(" ", dr("company"), dr("name"), dr("Code"))
        'Dim myView As Object
        myView = Draw.CreateDrawViewFromModelView3(dr("dz1") + dr("code") + dr("dz4"), "*前视", 0.115, 0.17, 0)
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager
        part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitSystem, 0, SwConst.swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(372, 204, 2)
        part.Extension.SetUserPreferenceInteger(516, 2, 0)
        part.Extension.SetUserPreferenceInteger(517, 208, 0)
        part.Extension.SetUserPreferenceInteger(372, 208, 2) '直径半径折弯标注
        设置图纸比例(dr("proportion"))
        part.ViewDisplayHidden '隐藏线不可见
        Dim SF As Object
        SF = part.Extension.InsertSurfaceFinishSymbol3(1, 0, 0, 0, 0, 0, 1, "", "", "", "", dr("roughness"), "", "") '插入右上角粗糙度
        SF.GetAnnotation().SetPosition2(190 / 1000, 275 / 1000, 0)

        Dim 分母 As Double
        Dim a As Object
        a = 索引字符串(dr("proportion"), ":")
        分母 = a(1)

        part.ActivateView("工程图视图1") '创建剪裁视图
        part.SketchManager.CreateCornerRectangle(0, 0, 0, -0.07 * 分母, 0.07 * 分母, 0)
        'Dim myView As Object
        myView = part.ActiveDrawingView
        myView.Crop()

        Dim myView1 As View
        part.ActivateView("工程图视图1")
        Dim L1s As SketchSegment
        L1s = part.SketchManager.CreateLine(0, 0, 0, 0, 0.07 * 分母, 0)
        L1s.Select4(False, Nothing)
        myView1 = part.CreateSectionViewAt5(0.115, 0.17, 0, " ", 128, Nothing, 0) '创建剖面图,并切换方向

        part.Extension.SelectByID2("工程图视图1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) '隐藏前视图
        part.SuppressView
        part.ClearSelection2(True)



        Object数组1 = 索引字符串（dr（"skills_requirement"））
        技术要求（0.031, 0.139, Object数组1（0）, Object数组1（1）, Object数组1（2））
        插入零件标号(myView1)

        插入BOM表(myView, 1)

        part.SaveAs3(dr("dz1") + dr("code") + dr("dz3"), 0, 8)



        dr.Close()
        mysqlcom.Dispose()
        mysqlcon.Close()
        mysqlcon.Dispose()



    End Sub

    Public Sub Exciter_rotor() '励磁机转子
        Dim mysqlcon As MySqlConnection
        Dim mysqlcom As MySqlCommand
        Dim dr As MySqlDataReader
        mysqlcon = New MySqlConnection("server=localhost" & ";userid=root" & ";password=123456" & ";database=eq_214_2431;pooling=false")
        '//打开数据库连接
        mysqlcon.Open()
        '//sql查询
        mysqlcom = New MySqlCommand("select * from _5eq_684_2042", mysqlcon)
        dr = mysqlcom.ExecuteReader()
        dr.Read()
        Do Until dr.GetString("id") = 1
            dr.Read()
        Loop



        Dim Code As String = dr("Code")
        Dim name As String = dr("name")
        Dim company As String = dr("company")
        Dim material As String = dr("material")


        swapp.OpenDoc6(dr("dz1") + dr("Asm" + 1.ToString + " Code") + dr("dz4"), 2, 32, "", errors, longwarnings)
        part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
        Component2 = part.AddComponent5(dr("dz1") + dr("Asm" + 1.ToString + " Code") + dr("dz4"), 0, "", False, "", 0, 0, 0)
        swapp.CloseDoc(dr("dz1") + dr("Asm" + 1.ToString + " Code") + dr("dz4"))



        Dim cusproper As SldWorks.CustomPropertyManager
        cusproper = part.Extension.CustomPropertyManager("")
        cusproper.Set2("名称", "励磁机转子")
        cusproper.Set2("代号", dr("code"))
        cusproper.Set2("材料", " ")


        part.SaveAs3(dr("dz1") + dr("code") + dr("dz4"), 0, 8)







        '工程图
        A4P(" ", dr("company"), dr("name"), dr("Code"))
        'Dim myView As Object
        myView = Draw.CreateDrawViewFromModelView3(dr("dz1") + dr("code") + dr("dz4"), "*前视", 0.115, 0.17, 0)
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager
        part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitSystem, 0, SwConst.swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(372, 204, 2)
        part.Extension.SetUserPreferenceInteger(516, 2, 0)
        part.Extension.SetUserPreferenceInteger(517, 208, 0)
        part.Extension.SetUserPreferenceInteger(372, 208, 2) '直径半径折弯标注
        设置图纸比例(dr("proportion"))
        'part.ViewDisplayHidden '隐藏线不可见


        Dim 分母 As Double
        Dim a As Object
        a = 索引字符串(dr("proportion"), ":")
        分母 = a(1)

        part.ActivateView("工程图视图1") '创建剪裁视图
        part.SketchManager.CreateCornerRectangle(0, 0, 0, -0.07 * 分母, 0.07 * 分母, 0)
        'Dim myView As Object
        myView = part.ActiveDrawingView
        myView.Crop()

        Dim myView1 As View
        part.ActivateView("工程图视图1")
        Dim L1s As SketchSegment
        L1s = part.SketchManager.CreateLine(0, 0, 0, 0, 0.07 * 分母, 0)
        L1s.Select4(False, Nothing)
        myView1 = part.CreateSectionViewAt5(0.115, 0.17, 0, " ", 128, Nothing, 0) '创建剖面图,并切换方向

        part.Extension.SelectByID2("工程图视图1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) '隐藏前视图
        part.SuppressView
        part.ClearSelection2(True)



        Object数组1 = 索引字符串（dr（"skills_requirement"））
        技术要求（0.031, 0.139, Object数组1（0）, Object数组1（1）, Object数组1（2）, Object数组1（3）, Object数组1（4）, Object数组1（5）, Object数组1（6））
        插入零件标号(myView1)


        插入BOM表(myView, 2)

        part.SaveAs3(dr("dz1") + dr("code") + dr("dz3"), 0, 8)

        dr.Close()
        mysqlcom.Dispose()
        mysqlcon.Close()
        mysqlcon.Dispose()
    End Sub

    Public Sub Exciter_rotor_assembly() '励磁机转子总成
        Dim mysqlcon As MySqlConnection
        Dim mysqlcom As MySqlCommand
        Dim dr As MySqlDataReader
        mysqlcon = New MySqlConnection("server=localhost" & ";userid=root" & ";password=123456" & ";database=eq_214_2431;pooling=false")
        '//打开数据库连接
        mysqlcon.Open()
        '//sql查询
        mysqlcom = New MySqlCommand("select * from _5eq_684_2041", mysqlcon)
        dr = mysqlcom.ExecuteReader()
        dr.Read()
        Do Until dr.GetString("id") = 1
            dr.Read()
        Loop



        Dim Code As String = dr("Code")
        Dim name As String = dr("name")
        Dim company As String = dr("company")
        Dim material As String = dr("material")


        Dim 小垫圈 As SldWorks.Component2
        Dim 弹簧垫圈 As SldWorks.Component2
        Dim 六角螺栓 As SldWorks.Component2

        swapp.OpenDoc6(dr("dz1") + dr("Asm" + 1.ToString + " Code") + dr("dz4"), 2, 32, "", errors, longwarnings)
        part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
        Component2 = part.AddComponent5(dr("dz1") + dr("Asm" + 1.ToString + " Code") + dr("dz4"), 0, "", False, "", 0, 0, 0)
        swapp.CloseDoc(dr("dz1") + dr("Asm" + 1.ToString + " Code") + dr("dz4"))

        For i = 3 To 5
            swapp.OpenDoc6(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0.05 * (i - 1))
            swapp.CloseDoc(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"))
            Select Case i > 0
                Case i = 3
                    弹簧垫圈 = Component2
                Case i = 4
                    小垫圈 = Component2
                Case i = 5
                    六角螺栓 = Component2
            End Select
        Next


        For i = 4 To 6
            Select Case i > 0

                Case i = 4
                    Component2 = 弹簧垫圈
                Case i = 5
                    Component2 = 小垫圈
                Case i = 6
                    Component2 = 六角螺栓
            End Select
            part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
            swbody = SelectionMgr.GetSelectedObject6(1, -1)
            Select Case i > 0
                Case i = 4
                    swbody4 = swbody
                Case i = 5
                    swbody5 = swbody
                Case i = 6
                    swbody6 = swbody


            End Select
            part.ClearSelection2(True)
        Next


        face4 = swbody4.GetFaces
        face5 = swbody5.GetFaces
        face6 = swbody6.GetFaces

        edge4 = swbody4.GetEdges
        edge5 = swbody5.GetEdges
        edge6 = swbody6.GetEdges



        Dim body1() As Object
        part.Extension.SelectByID2(dr("Asm1 Code") + "-1@*/" + dr("A1A1 Code") + "-1@" + dr("Asm1 Code") + "/" + dr("A1A1P1 Code") + "-1@" + dr("A1A1 Code"), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)

        Component2 = SelectionMgr.GetSelectedObject6(1, -1)
        body1 = Component2.GetBodies3(0, 0)
        face1 = body1(0).GetFaces
        edge1 = body1(0).GetEdges


        '遍历（1, edge6, 400, 0）


        SelectionMgr.AddSelectionListObject(edge6(0), Nothing)
        SelectionMgr.AddSelectionListObject(edge5(0), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face6(0), Nothing)
        SelectionMgr.AddSelectionListObject(face5(2), Nothing)
        配合（"重合", 1）

        SelectionMgr.AddSelectionListObject(edge6(0), Nothing)
        SelectionMgr.AddSelectionListObject(edge4(4), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face4(2), Nothing)
        SelectionMgr.AddSelectionListObject(face5(1), Nothing)
        配合（"重合", 1）

        SelectionMgr.AddSelectionListObject(edge6(0), Nothing)
        SelectionMgr.AddSelectionListObject(edge1(177), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face4(1), Nothing)
        SelectionMgr.AddSelectionListObject(face1(17), Nothing)
        配合（"重合", 1）



        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(edge1(59）, Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        小垫圈.Select2(True, 1)
        弹簧垫圈.Select2(True, 1)
        六角螺栓.Select2(True, 1)
        part.FeatureManager.FeatureCircularPattern5(2, 2 * PI, False, "NULL", False, True, False, False, False, False, 1, 0, "NULL", False)



        part.ClearSelection2(True)
        小垫圈.Select2(False, 1)
        弹簧垫圈.Select2(True, 1)
        六角螺栓.Select2(True, 1)
        part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, True, 2, Nothing, 0) '镜像零件
        Dim compsToMirror(2) As SldWorks.Component2
        Dim swMirrorPlane As SldWorks.Feature
        compsToMirror(0) = SelectionMgr.GetSelectedObjectsComponent4(1, 1)
        compsToMirror(1) = SelectionMgr.GetSelectedObjectsComponent4(2, 1)
        compsToMirror(2) = SelectionMgr.GetSelectedObjectsComponent4(3, 1)
        swMirrorPlane = SelectionMgr.GetSelectedObject6(1, 2)
        AssemblyDoc.MirrorComponents3(swMirrorPlane, Nothing, 1, True, compsToMirror, True, "", 2, Nothing, "", 1, False, True, False)

        part.ClearSelection2(True)
        小垫圈.Select2(False, 1)
        弹簧垫圈.Select2(True, 1)
        六角螺栓.Select2(True, 1)
        part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 2, Nothing, 0) '镜像零件
        'Dim compsToMirror(2) As SldWorks.Component2
        'Dim swMirrorPlane As SldWorks.Feature
        compsToMirror(0) = SelectionMgr.GetSelectedObjectsComponent4(1, 1)
        compsToMirror(1) = SelectionMgr.GetSelectedObjectsComponent4(2, 1)
        compsToMirror(2) = SelectionMgr.GetSelectedObjectsComponent4(3, 1)
        swMirrorPlane = SelectionMgr.GetSelectedObject6(1, 2)
        AssemblyDoc.MirrorComponents3(swMirrorPlane, Nothing, 1, True, compsToMirror, True, "", 2, Nothing, "", 1, False, True, False)
        part.ClearSelection2(True)


        Dim cusproper As SldWorks.CustomPropertyManager
        cusproper = part.Extension.CustomPropertyManager("")
        cusproper.Set2("名称", "励磁机转子总成")
        cusproper.Set2("代号", dr("code"))
        cusproper.Set2("材料", " ")
        part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True) '隐藏标注

        part.SaveAs3(dr("dz1") + dr("code") + dr("dz4"), 0, 8)






        '工程图

        A3模板(" ", dr("company"), dr("name"), dr("Code"))
        myView = Draw.CreateDrawViewFromModelView3(dr("dz1") + dr("code") + dr("dz4"), "*前视", 0.12, 0.18, 0)
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager
        part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitSystem, 0, SwConst.swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(372, 204, 2)
        part.Extension.SetUserPreferenceInteger(516, 2, 0)
        part.Extension.SetUserPreferenceInteger(517, 208, 0)
        part.Extension.SetUserPreferenceInteger(372, 208, 2) '直径半径折弯标注
        设置图纸比例(dr("proportion"))
        part.ViewDisplayHidden '隐藏线不可见


        Dim 分母 As Double
        Dim a As Object
        a = 索引字符串(dr("proportion"), ":")
        分母 = a(1)

        Initial_setting(1) '打开捕捉
        Dim myView1 As View
        part.ActivateView("工程图视图1")
        Dim L1s, l2s, l3s As SketchSegment
        L1s = part.SketchManager.CreateLine(0, 0, 0, 0, 0.07 * 分母, 0)
        l2s = part.SketchManager.CreateLine(0, 0, 0, -0.07 * 分母, 0, 0)
        l3s = part.SketchManager.CreateLine(-0.07 * 分母, 0, 0, -0.07 * 分母, -0.07 * 分母, 0)
        L1s.Select4(False, Nothing)
        l2s.Select4(True, Nothing)
        l3s.Select4(True, Nothing)
        myView1 = part.CreateSectionViewAt5(0.3, 0, 0, " ", 132, Nothing, 0) '创建剖面图,并切换方向
        Initial_setting(0) '关闭捕捉

        Object数组1 = 索引字符串（dr（"skills_requirement"））
        技术要求（0.08, 0.09, Object数组1（0）, Object数组1（1）, Object数组1（2）, Object数组1（3））
        插入零件标号(myView1)



        插入BOM表(myView, 2)
        part.SaveAs3(dr("dz1") + dr("code") + dr("dz3"), 0, 8)

        dr.Close()
        mysqlcom.Dispose()
        mysqlcon.Close()
        mysqlcon.Dispose()
    End Sub


    Public Sub Rotor() '转子
        Dim mysqlcon As MySqlConnection
        Dim mysqlcom As MySqlCommand
        Dim dr As MySqlDataReader
        mysqlcon = New MySqlConnection("server=localhost" & ";userid=root" & ";password=123456" & ";database=eq_214_2431;pooling=false")
        '//打开数据库连接
        mysqlcon.Open()
        '//sql查询
        mysqlcom = New MySqlCommand("select * from _5eq_675_2219", mysqlcon)
        dr = mysqlcom.ExecuteReader()
        dr.Read()
        Do Until dr.GetString("id") = 1
            dr.Read()
        Loop



        Dim Code As String = dr("Code")
        Dim name As String = dr("name")
        Dim company As String = dr("company")
        Dim material As String = dr("material")


        '第一段,轴套部分配合
        Dim 转轴 As SldWorks.Component2
        Dim 连接轴片 As SldWorks.Component2
        Dim 轴套 As SldWorks.Component2


        'swapp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUseFolderAsDefaultSearchLocation, False) '设置异形孔向导/toolbox取消打勾

        For i = 1 To 3
            swapp.OpenDoc6(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0.05 * (i - 1))
            swapp.CloseDoc(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"))
            Select Case i > 0
                Case i = 1
                    转轴 = Component2
                Case i = 2
                    连接轴片 = Component2
                Case i = 3
                    轴套 = Component2
            End Select
        Next


        For i = 1 To 3
            Select Case i > 0

                Case i = 1
                    Component2 = 转轴
                Case i = 2
                    Component2 = 连接轴片
                Case i = 3
                    Component2 = 轴套
            End Select
            part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
            swbody = SelectionMgr.GetSelectedObject6(1, -1)
            Select Case i > 0
                Case i = 1
                    swbody1 = swbody
                Case i = 2
                    swbody2 = swbody
                Case i = 3
                    swbody3 = swbody


            End Select
            part.ClearSelection2(True)
        Next


        face1 = swbody1.GetFaces
        face2 = swbody2.GetFaces
        face3 = swbody3.GetFaces

        edge1 = swbody1.GetEdges
        edge2 = swbody2.GetEdges
        edge3 = swbody3.GetEdges


        '遍历（1, face1, 400, 0）

        Dim 螺栓M16 As SldWorks.Component2
        Dim 平键22 As SldWorks.Component2
        For i = 4 To 5
            swapp.OpenDoc6(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0.05 * (i - 1))
            swapp.CloseDoc(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"))
            Select Case i > 0
                Case i = 4
                    螺栓M16 = Component2
                Case i = 5
                    平键22 = Component2
            End Select
        Next


        For i = 4 To 5
            Select Case i > 0

                Case i = 4
                    Component2 = 螺栓M16
                Case i = 5
                    Component2 = 平键22

            End Select
            part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
            swbody = SelectionMgr.GetSelectedObject6(1, -1)
            Select Case i > 0
                Case i = 4
                    swbody4 = swbody
                Case i = 5
                    swbody5 = swbody


            End Select
            part.ClearSelection2(True)
        Next

        face4 = swbody4.GetFaces
        face5 = swbody5.GetFaces

        edge4 = swbody4.GetEdges
        edge5 = swbody5.GetEdges


        '键配合
        SelectionMgr.AddSelectionListObject(face1(38), Nothing)
        SelectionMgr.AddSelectionListObject(face5(0), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face1(39), Nothing)
        SelectionMgr.AddSelectionListObject(face5(1), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(face1(42), Nothing)
        SelectionMgr.AddSelectionListObject(face5(5), Nothing)
        配合（"重合", 1）

        '轴套配合
        SelectionMgr.AddSelectionListObject(face1(24), Nothing)
        SelectionMgr.AddSelectionListObject(face3(0), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(face1(3), Nothing)
        SelectionMgr.AddSelectionListObject(face3(2), Nothing)
        配合（"同轴", 1）
        SelectionMgr.AddSelectionListObject(face3(9), Nothing)
        SelectionMgr.AddSelectionListObject(face5(4), Nothing)
        配合（"平行", 1）

        '连轴片
        SelectionMgr.AddSelectionListObject(face1(3), Nothing)
        SelectionMgr.AddSelectionListObject(face2(3), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face2(1), Nothing)
        SelectionMgr.AddSelectionListObject(face3(1), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(face3(11), Nothing)
        SelectionMgr.AddSelectionListObject(face2(5), Nothing)
        配合（"同轴", 0）

        '螺丝
        SelectionMgr.AddSelectionListObject(face3(11), Nothing)
        SelectionMgr.AddSelectionListObject(face4(1), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face2(2), Nothing)
        SelectionMgr.AddSelectionListObject(face4(0), Nothing)
        配合（"重合", 1）

        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(face1(3）, Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        螺栓M16.Select2(True, 1)
        part.FeatureManager.FeatureCircularPattern5(8, 2 * PI, False, "NULL", False, True, False, False, False, False, 1, 0, "NULL", False) '圆周阵列




        '第二段,风扇部分配合

        Dim 风扇 As SldWorks.Component2

        For i = 6 To 6
            swapp.OpenDoc6(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0.05 * (i - 1))
            swapp.CloseDoc(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"))
            风扇 = Component2
        Next

        part.Extension.SelectByID2(风扇.IGetBody.Name + "@" + 风扇.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
        swbody = SelectionMgr.GetSelectedObject6(1, -1)
        swbody2 = swbody

        face2 = swbody2.GetFaces
        edge2 = swbody2.GetEdges



        Dim 螺栓M10 As SldWorks.Component2
        Dim 螺母M10 As SldWorks.Component2
        Dim 挡圈 As SldWorks.Component2
        Dim 平键18 As SldWorks.Component2
        For i = 9 To 10
            swapp.OpenDoc6(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0.05 * (i - 1))
            swapp.CloseDoc(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"))
            Select Case i > 0
                Case i = 7
                    螺栓M10 = Component2
                Case i = 8
                    螺母M10 = Component2
                Case i = 9
                    挡圈 = Component2
                Case i = 10
                    平键18 = Component2
            End Select
        Next


        For i = 5 To 6
            Select Case i > 0

                Case i = 3
                    Component2 = 螺栓M10
                Case i = 4
                    Component2 = 螺母M10
                Case i = 5
                    Component2 = 挡圈
                Case i = 6
                    Component2 = 平键18
            End Select
            part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
            swbody = SelectionMgr.GetSelectedObject6(1, -1)
            Select Case i > 0

                Case i = 3
                    swbody3 = swbody
                Case i = 4
                    swbody4 = swbody
                Case i = 5
                    swbody5 = swbody
                Case i = 6
                    swbody6 = swbody
            End Select
            part.ClearSelection2(True)
        Next

        face3 = swbody3.GetFaces
        face4 = swbody4.GetFaces
        face5 = swbody5.GetFaces
        face6 = swbody6.GetFaces

        edge3 = swbody3.GetEdges
        edge4 = swbody4.GetEdges
        edge5 = swbody5.GetEdges
        edge6 = swbody6.GetEdges

        '键配合
        SelectionMgr.AddSelectionListObject(face1(45), Nothing)
        SelectionMgr.AddSelectionListObject(face6(0), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face1(46), Nothing)
        SelectionMgr.AddSelectionListObject(face6(1), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(face1(47), Nothing)
        SelectionMgr.AddSelectionListObject(face6(5), Nothing)
        配合（"重合", 1）

        '挡圈配合
        SelectionMgr.AddSelectionListObject(face1(2), Nothing)
        SelectionMgr.AddSelectionListObject(face5(2), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face1(16), Nothing)
        SelectionMgr.AddSelectionListObject(face5(9), Nothing)
        配合（"重合", 1）

        '风扇装配
        SelectionMgr.AddSelectionListObject(face1(2), Nothing)
        SelectionMgr.AddSelectionListObject(face2(10), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face2(9), Nothing)
        SelectionMgr.AddSelectionListObject(face5(8), Nothing)
        配合（"重合", 1）


        '遍历（1, face2, 400, 0）

        '带绕组转子铁芯配合


        '键配合
        swapp.OpenDoc6(dr("Standard Parts address") + dr("Part" + 11.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
        part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
        Component2 = part.AddComponent5(dr("Standard Parts address") + dr("Part" + 11.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0.05 * 4)
        swapp.CloseDoc(dr("Standard Parts address") + dr("Part" + 11.ToString + " Code") + dr("dz2"))
        part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
        swbody = SelectionMgr.GetSelectedObject6(1, -1)
        swbody3 = swbody
        face3 = swbody3.GetFaces
        edge3 = swbody3.GetEdges

        SelectionMgr.AddSelectionListObject(face1(50), Nothing)
        SelectionMgr.AddSelectionListObject(face3(0), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face1(51), Nothing)
        SelectionMgr.AddSelectionListObject(face3(1), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(face1(52), Nothing)
        SelectionMgr.AddSelectionListObject(face3(5), Nothing)
        配合（"重合", 1）






        '转子铁芯与轴配合


        swapp.OpenDoc6(dr("dz1") + dr("Asm" + 1.ToString + " Code") + dr("dz4"), 2, 32, "", errors, longwarnings)
        part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
        Component2 = part.AddComponent5(dr("dz1") + dr("Asm" + 1.ToString + " Code") + dr("dz4"), 0, "", False, "", 0, 0, 0)
        swapp.CloseDoc(dr("dz1") + dr("Asm" + 1.ToString + " Code") + dr("dz4"))

        Dim body2(), body3() As Object
        '冲片
        part.Extension.SelectByID2(dr("Asm1 Code") + "-1@*/" + dr("A1_1 Code") + "-1@" + dr("Asm1 Code") + "/" + dr("A1_1P1 Code") + "-1@" + dr("A1_1 Code"), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        'part.Extension.SelectByID2("5EQ.526.2137-1@装配体38/5EQ.676.2105-1@5EQ.526.2137/8EQ.630.2005-1@5EQ.676.2105", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Component2 = SelectionMgr.GetSelectedObject6(1, -1)
        body2 = Component2.GetBodies3(0, 0)
        face2 = body2(0).GetFaces
        edge2 = body2(0).GetEdges

        SelectionMgr.AddSelectionListObject(face1(1), Nothing)
        SelectionMgr.AddSelectionListObject(face2(50), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face2(51), Nothing)
        SelectionMgr.AddSelectionListObject(face3(3), Nothing)
        配合（"重合", 1）


        '阻尼板
        part.Extension.SelectByID2(dr("Asm1 Code") + "-1@*/" + dr("A1_1 Code") + "-1@" + dr("Asm1 Code") + "/" + dr("A1_1P2 Code") + "-2@" + dr("A1_1 Code"), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        'part.Extension.SelectByID2("5EQ.526.2137-1@装配体38/5EQ.676.2105-1@5EQ.526.2137/8EQ.548.2109-2@5EQ.676.2105", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Component2 = SelectionMgr.GetSelectedObject6(1, -1)
        body3 = Component2.GetBodies3(0, 0)
        face3 = body3(0).GetFaces
        edge3 = body3(0).GetEdges

        SelectionMgr.AddSelectionListObject(face1(12), Nothing)
        SelectionMgr.AddSelectionListObject(face3(1), Nothing)
        配合（"重合", 1）

        '定位棒
        part.Extension.SelectByID2(dr("Asm1 Code") + "-1@*/" + dr("A1_1 Code") + "-1@" + dr("Asm1 Code") + "/" + dr("A1_1P3 Code") + "-1@" + dr("A1_1 Code"), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        'Part.Extension.SelectByID2("5EQ.526.2137-1@装配体7/5EQ.676.2105-1@5EQ.526.2137/8EQ.272.2181-4@5EQ.676.2105", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Component2 = SelectionMgr.GetSelectedObject6(1, -1)
        body2 = Component2.GetBodies3(0, 0)
        face2 = body2(0).GetFaces
        edge2 = body2(0).GetEdges


        Dim 平衡环 As SldWorks.Component2
        Dim 平衡铁 As SldWorks.Component2

        For i = 12 To 13
            swapp.OpenDoc6(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0 + (i - 1) * 0.05)
            swapp.CloseDoc(dr("Part address") + dr("Part" + i.ToString + " Code") + dr("dz2"))

            Select Case i > 0
                Case i = 12
                    平衡环 = Component2
                Case i = 13
                    平衡铁 = Component2

            End Select
        Next

        For i = 1 To 2
            Select Case i > 0
                Case i = 1
                    Component2 = 平衡环
                Case i = 2
                    Component2 = 平衡铁

            End Select
            part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
            swbody = SelectionMgr.GetSelectedObject6(1, -1)
            Select Case i > 0
                Case i = 1
                    swbody3 = swbody
                Case i = 2
                    swbody4 = swbody
            End Select
            part.ClearSelection2(True)
        Next

        Dim 螺栓M6 As SldWorks.Component2
        Dim 螺母M6 As SldWorks.Component2
        Dim 垫圈6 As SldWorks.Component2
        Dim 螺栓M8 As SldWorks.Component2
        Dim 垫圈8 As SldWorks.Component2

        For i = 14 To 18
            swapp.OpenDoc6(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0.05 * (i - 1))
            swapp.CloseDoc(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"))
            Select Case i > 0
                Case i = 14
                    螺母M6 = Component2
                Case i = 15
                    螺栓M6 = Component2
                Case i = 16
                    垫圈6 = Component2
                Case i = 17
                    螺栓M8 = Component2
                Case i = 18
                    垫圈8 = Component2
            End Select
        Next


        For i = 3 To 7
            Select Case i > 0

                Case i = 3
                    Component2 = 螺母M6
                Case i = 4
                    Component2 = 螺栓M6
                Case i = 5
                    Component2 = 垫圈6
                Case i = 6
                    Component2 = 螺栓M8
                Case i = 7
                    Component2 = 垫圈8
            End Select
            part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
            swbody = SelectionMgr.GetSelectedObject6(1, -1)
            Select Case i > 0

                Case i = 3
                    swbody5 = swbody
                Case i = 4
                    swbody6 = swbody
                Case i = 5
                    swbody7 = swbody
                Case i = 6
                    swbody8 = swbody
                Case i = 7
                    swbody9 = swbody
            End Select
            part.ClearSelection2(True)
        Next

        face3 = swbody3.GetFaces
        face4 = swbody4.GetFaces
        face5 = swbody5.GetFaces
        face6 = swbody6.GetFaces
        face7 = swbody7.GetFaces
        face8 = swbody8.GetFaces
        face9 = swbody9.GetFaces

        edge3 = swbody3.GetEdges
        edge4 = swbody4.GetEdges
        edge5 = swbody5.GetEdges
        edge6 = swbody6.GetEdges
        edge7 = swbody7.GetEdges
        edge8 = swbody8.GetEdges
        edge9 = swbody9.GetEdges

        '遍历（1, face3, 400, 0）
        '平衡环
        SelectionMgr.AddSelectionListObject(face1(1), Nothing)
        SelectionMgr.AddSelectionListObject(face3(3), Nothing)
        配合（"同轴", 1）
        SelectionMgr.AddSelectionListObject(face2(6), Nothing)
        SelectionMgr.AddSelectionListObject(face3(5), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face2(2), Nothing)
        SelectionMgr.AddSelectionListObject(face3(1), Nothing)
        配合（"重合", 1）
        '弹簧垫圈8
        SelectionMgr.AddSelectionListObject(face9(3), Nothing)
        SelectionMgr.AddSelectionListObject(face3(5), Nothing)
        配合（"同轴", 1）
        SelectionMgr.AddSelectionListObject(face9(2), Nothing)
        SelectionMgr.AddSelectionListObject(face3(2), Nothing)
        配合（"重合", 1）
        '螺栓8
        SelectionMgr.AddSelectionListObject(face8(21), Nothing)
        SelectionMgr.AddSelectionListObject(face3(5), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face9(1), Nothing)
        SelectionMgr.AddSelectionListObject(face8(0), Nothing)
        配合（"重合", 1）

        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(face1(3）, Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        螺栓M8.Select2(True, 1)
        垫圈8.Select2(True, 1)
        part.FeatureManager.FeatureCircularPattern5(4, 2 * PI, False, "NULL", False, True, False, False, False, False, 1, 0, "NULL", False) '圆周阵列



        '平衡铁
        SelectionMgr.AddSelectionListObject(face4(3), Nothing)
        SelectionMgr.AddSelectionListObject(face3(10), Nothing)
        配合（"同轴", 1）
        SelectionMgr.AddSelectionListObject(face4(2), Nothing)
        SelectionMgr.AddSelectionListObject(face3(2), Nothing)
        配合（"重合", 1）
        '弹簧垫圈6
        SelectionMgr.AddSelectionListObject(face7(3), Nothing)
        SelectionMgr.AddSelectionListObject(face4(3), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face7(2), Nothing)
        SelectionMgr.AddSelectionListObject(face4(1), Nothing)
        配合（"重合", 1）
        '螺栓6
        SelectionMgr.AddSelectionListObject(face6(19), Nothing)
        SelectionMgr.AddSelectionListObject(face3(10), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face6(0), Nothing)
        SelectionMgr.AddSelectionListObject(face7(1), Nothing)
        配合（"重合", 1）
        '螺母6
        SelectionMgr.AddSelectionListObject(face5(6), Nothing)
        SelectionMgr.AddSelectionListObject(face3(1), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(face5(8), Nothing)
        SelectionMgr.AddSelectionListObject(face3(10), Nothing)
        配合（"同轴", 1）

        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(face1(3）, Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        螺栓M6.Select2(True, 1)
        垫圈6.Select2(True, 1)
        螺母M6.Select2(True, 1)
        平衡铁.Select2(True, 1)
        Feature = part.FeatureManager.FeatureCircularPattern5(2, -30 * PI / 180, False, "NULL", False, False, False, False, False, False, 1, 0, "NULL", False) '圆周阵列

        part.ClearSelection2(True)
        SelectionMgr.AddSelectionListObject(face1(3）, Nothing)
        SelectionMgr.SetSelectedObjectMark(1, 2, 0) '设定标记
        螺栓M6.Select2(True, 1)
        垫圈6.Select2(True, 1)
        螺母M6.Select2(True, 1)
        平衡铁.Select2(True, 1)
        Feature.Select2(True, 1)
        part.FeatureManager.FeatureCircularPattern5(4, 2 * PI, False, "NULL", False, True, False, False, False, False, 1, 0, "NULL", False) '圆周阵列
        '遍历（1, face5, 400, 0）









        '励磁机转子总成与轴配合

        For i = 19 To 20
            swapp.OpenDoc6(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 1, 32, "", errors, longwarnings)
            part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
            Component2 = part.AddComponent5(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"), 0, "", False, "", 0, 0, 0.05 * (i - 1))
            swapp.CloseDoc(dr("Standard Parts address") + dr("Part" + i.ToString + " Code") + dr("dz2"))
            Select Case i > 0
                Case i = 19
                    平键18 = Component2
                Case i = 20
                    挡圈 = Component2
            End Select
        Next


        For i = 5 To 6
            Select Case i > 0
                Case i = 5
                    Component2 = 平键18
                Case i = 6
                    Component2 = 挡圈
            End Select
            part.Extension.SelectByID2(Component2.IGetBody.Name + "@" + Component2.Name2() + "@" + AssemblyTitle, "SOLIDBODY", 0, 0, 0, False, 1, Nothing, 0)
            swbody = SelectionMgr.GetSelectedObject6(1, -1)
            Select Case i > 0
                Case i = 5
                    swbody3 = swbody
                Case i = 6
                    swbody4 = swbody
            End Select
            part.ClearSelection2(True)
        Next

        face3 = swbody3.GetFaces
        face4 = swbody4.GetFaces

        edge3 = swbody3.GetEdges
        edge4 = swbody4.GetEdges

        '键配合
        SelectionMgr.AddSelectionListObject(face1(55), Nothing)
        SelectionMgr.AddSelectionListObject(face3(0), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face1(56), Nothing)
        SelectionMgr.AddSelectionListObject(face3(1), Nothing)
        配合（"平行", 1）
        SelectionMgr.AddSelectionListObject(face1(57), Nothing)
        SelectionMgr.AddSelectionListObject(face3(5), Nothing)
        配合（"重合", 1）

        '挡圈配合
        SelectionMgr.AddSelectionListObject(face1(0), Nothing)
        SelectionMgr.AddSelectionListObject(face4(2), Nothing)
        配合（"同轴", 0）
        SelectionMgr.AddSelectionListObject(face1(27), Nothing)
        SelectionMgr.AddSelectionListObject(face4(8), Nothing)
        配合（"重合", 1）

        '励磁机转子总成
        swapp.OpenDoc6(dr("dz1") + dr("Asm" + 2.ToString + " Code") + dr("dz4"), 2, 32, "", errors, longwarnings)
        part = swapp.ActivateDoc3(AssemblyTitle, True, 0, errors)
        Component2 = part.AddComponent5(dr("dz1") + dr("Asm" + 2.ToString + " Code") + dr("dz4"), 0, "", False, "", 0, 0, 0.1)
        swapp.CloseDoc(dr("dz1") + dr("Asm" + 2.ToString + " Code") + dr("dz4"))

        '遍历（1, face4, 400, 0）

        part.Extension.SelectByID2(dr("Asm2 Code") + "-1@*/" + dr("A2_1 Code") + "-1@" + dr("Asm2 Code") + "/" + dr("A2_1_1 Code") + "-1@" + dr("A2_1 Code") + "/" + dr("A2_1_1P1 Code") + "-1@" + dr("A2_1_1 Code"), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        'Part.Extension.SelectByID2("5EQ.684.2041-1@装配体54/5EQ.684.2042-1@5EQ.684.2041/5EQ.682.2018-1@5EQ.684.2042/8EQ.150.1005-1@5EQ.682.2018", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Component2 = SelectionMgr.GetSelectedObject6(1, -1)
        body2 = Component2.GetBodies3(0, 0)
        face2 = body2(0).GetFaces
        edge2 = body2(0).GetEdges

        SelectionMgr.AddSelectionListObject(face2(33), Nothing)
        SelectionMgr.AddSelectionListObject(face1(0), Nothing)
        配合（"同轴", 1）
        SelectionMgr.AddSelectionListObject(face2(29), Nothing)
        SelectionMgr.AddSelectionListObject(face1(6), Nothing)
        配合（"重合", 1）
        SelectionMgr.AddSelectionListObject(face2(34), Nothing)
        SelectionMgr.AddSelectionListObject(face3(1), Nothing)
        配合（"平行", 1）

        ''遍历（1, face2, 400, 0）
        Dim cusproper As SldWorks.CustomPropertyManager
        cusproper = part.Extension.CustomPropertyManager("")
        cusproper.Set2("名称", "转子")
        cusproper.Set2("代号", dr("code"))
        cusproper.Set2("材料", " ")
        part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True) '隐藏标注
        part.Extension.SelectByID2("LimitDistance1", "MATE", 0, 0, 0, False, 0, Nothing, 0)
        part.EditDelete()
        part.SaveAs3(dr("dz1") + dr("code") + dr("dz4"), 0, 8)

    End Sub





    Public Sub Rotor_drw()
        Dim mysqlcon As MySqlConnection
        Dim mysqlcom As MySqlCommand
        Dim dr As MySqlDataReader
        mysqlcon = New MySqlConnection("server=localhost" & ";userid=root" & ";password=123456" & ";database=eq_214_2431;pooling=false")
        '//打开数据库连接
        mysqlcon.Open()
        '//sql查询
        mysqlcom = New MySqlCommand("select * from _5eq_675_2219", mysqlcon)
        dr = mysqlcom.ExecuteReader()
        dr.Read()
        Do Until dr.GetString("id") = 1
            dr.Read()
        Loop


        '工程图

        A3模板(" ", dr("company"), dr("name"), dr("Code"))
        myView = Draw.CreateDrawViewFromModelView3(dr("dz1") + dr("code") + dr("dz4"), "*右视", 0.21, 0.19, 0)
        part = swapp.ActiveDoc
        Draw = swapp.ActiveDoc
        SelectionMgr = part.SelectionManager
        part.Extension.SetUserPreferenceInteger(SwConst.swUserPreferenceIntegerValue_e.swUnitSystem, 0, SwConst.swUnitSystem_e.swUnitSystem_MMGS)
        part.Extension.SetUserPreferenceInteger(372, 204, 2)
        part.Extension.SetUserPreferenceInteger(516, 2, 0)
        part.Extension.SetUserPreferenceInteger(517, 208, 0)
        part.Extension.SetUserPreferenceInteger(372, 208, 2) '直径半径折弯标注
        设置图纸比例(dr("proportion"))
        part.ViewDisplayHidden '隐藏线不可见
        part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True) '不显示草图

        part.ClearSelection2(True)

        Dim 分母 As Double
        Dim a As Object
        a = 索引字符串(dr("proportion"), ":")
        分母 = a(1)


        part.ActivateView("工程图视图1")
        part.SketchManager.CreateCornerRectangle(0, 0, 0, -0.09 * 分母, 0.09 * 分母, 0)
        myView = part.ActiveDrawingView
        myView.Crop()

        Dim myView1 As View
        part.ActivateView("工程图视图1")
        Dim L1s As SketchSegment
        L1s = part.SketchManager.CreateLine(0, 0, 0, 0, 0.09 * 分母, 0)
        L1s.Select4(False, Nothing)
        myView1 = part.CreateSectionViewAt5(0.21, 0.21, 0, " ", 128, Nothing, 0) '创建剖面图,并切换方向

        part.Extension.SelectByID2("工程图视图1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) '隐藏视图
        part.SuppressView
        part.ClearSelection2(True)



        Object数组1 = 索引字符串（dr（"skills_requirement"））
        技术要求（0.08, 0.15, Object数组1（0）, Object数组1（1）, Object数组1（2）, Object数组1（3）, Object数组1（4）, Object数组1（5））
        插入零件标号(myView1)

        插入BOM表(myView1, 2)
        part.SaveAs3(dr("dz1") + dr("code") + dr("dz3"), 0, 8)


        dr.Close()
        mysqlcom.Dispose()
        mysqlcon.Close()
        mysqlcon.Dispose()


    End Sub

    Public Sub 配合(类型$, Optional 反向% = 0, Optional 距离# = 0, Optional 距离翻转 As Boolean = False, Optional 距离上限# = 1, Optional 距离下限# = 0)
        If 类型 = "重合" Then
            AssemblyDoc.AddMate5(0, 反向, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, 0) '重合
        ElseIf 类型 = "同轴" Then
            AssemblyDoc.AddMate5(1, 反向, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, 0) '同轴
        ElseIf 类型 = "垂直" Then
            AssemblyDoc.AddMate5(2, 反向, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, 0) '垂直
        ElseIf 类型 = "平行" Then
            AssemblyDoc.AddMate5(3, 反向, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, 0) '平行
        ElseIf 类型 = "相切" Then
            AssemblyDoc.AddMate5(4, 反向, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, 0) '相切
        ElseIf 类型 = "距离" Then
            AssemblyDoc.AddMate5(5, 反向, 距离翻转, 距离, 距离上限, 距离下限, 0, 0, 0, 0, 0, False, False, 0, 0) '距离
        End If
        part.ClearSelection2(True)
        part.EditRebuild3()
    End Sub
    Public Sub 渲染(entity As Object, 外观路径$)

        Dim displayStateNames As Object

        Dim status As Boolean

        Dim modelName As String

        Dim materialName As String

        Dim errors As Integer

        Dim warnings As Integer

        Dim nbrDisplayStates As Integer

        Dim i As Integer

        Dim k As Integer

        Dim nbrMaterials As Integer

        Dim materialID1 As Integer

        Dim materialID2 As Integer

        Dim materialID1_ToDelete(0) As Integer

        Dim materialID2_ToDelete(0) As Integer



        ' Get active configuration and create a new display

        ' state for this configuration

        swConfig = part.GetActiveConfiguration

        status = swConfig.CreateDisplayState("Display State 2")

        part.ForceRebuild3(True)

        ' Get active configuration and create another new

        ' display state for this configuration

        swConfig = part.GetActiveConfiguration

        status = swConfig.CreateDisplayState("Display State 3")

        part.ForceRebuild3(True)

        ' Create appearance

        materialName = 外观路径

        RenderMaterial = part.Extension.CreateRenderMaterial(materialName)


        status = entity.Select2(False, 1)
        status = RenderMaterial.AddEntity(entity)

        ' Get the names of display states

        displayStateNames = swConfig.GetDisplayStates

        nbrDisplayStates = swConfig.GetDisplayStatesCount

        Debug.Print("This configuration's display states =")

        For i = 0 To (nbrDisplayStates - 1)

            Debug.Print(" Display state name = " & displayStateNames(i))

        Next i

        ' Add appearance to all of the display states

        status = part.Extension.AddDisplayStateSpecificRenderMaterial(RenderMaterial, SwConst.swDisplayStateOpts_e.swAllDisplayState, displayStateNames, materialID1, materialID2)




    End Sub
    Public Sub 遍历(同时遍历多实体数%, 实体集合() As Object, 上限#, Optional 初值% = 0)
        part.ClearSelection2(True)
        Dim vEdgeCount As Integer
        vEdgeCount = 初值
        Do Until vEdgeCount >= 上限
            For i = vEdgeCount To vEdgeCount + 同时遍历多实体数 - 1
                SelectionMgr.AddSelectionListObject(实体集合(i), Nothing)
            Next
            vEdgeCount = vEdgeCount + 同时遍历多实体数
            MsgBox(vEdgeCount - 1)
            part.ClearSelection2(True)
        Loop
    End Sub

    Public Function 基准轴(基准$) As SldWorks.Feature
        Dim AXIS As SldWorks.Feature
        Dim line1 As SldWorks.SketchLine
        Dim line1Segment As SldWorks.SketchSegment
        If 基准 = "Z" Then
            part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            part.SketchManager.InsertSketch(True)
            line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, -0.01, 0)
            line1Segment = line1
            part.InsertSketch2(True)
            line1Segment.Select4(False, Nothing)
            part.InsertAxis2(True)
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            AXIS.Select2(False, Nothing)
        End If
        If 基准 = "Y" Then
            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            part.SketchManager.InsertSketch(True)
            line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.01, 0)
            line1Segment = line1
            part.InsertSketch2(True)
            line1Segment.Select4(False, Nothing)
            part.InsertAxis2(True)
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            AXIS.Select2(False, Nothing)
        End If
        If 基准 = "X" Then
            part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            part.SketchManager.InsertSketch(True)
            line1 = part.SketchManager.CreateCenterLine(0, 0, 0, 0.01, 0, 0)
            line1Segment = line1
            part.InsertSketch2(True)
            line1Segment.Select4(False, Nothing)
            part.InsertAxis2(True)
            AXIS = SelectionMgr.GetSelectedObject6(1, -1)
            AXIS.Select2(False, Nothing)
        End If
        AXIS = SelectionMgr.GetSelectedObject6(1, -1)
        基准轴 = AXIS
    End Function
    Public Sub A4P(material$, company$, name$, Code$)
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a4p.drwdot", 1, 0, 0)
        part = swapp.ActiveDoc
        myNote = part.InsertNote("公司名称")
        If Not myNote Is Nothing Then
            myNote.LockPosition = False
            myNote.Angle = 0
            myNote.SetBalloon(0, 0)
            Annotation = myNote.GetAnnotation()
            If Not Annotation Is Nothing Then
                Annotation.SetLeader3(SwConst.swLeaderStyle_e.swNO_LEADER, 0, True, False, False, False)
                Annotation.SetPosition(0.164493037612998, 0.0537752937136617, 0)
                Annotation.SetTextFormat(0, True, 0)
            End If
        End If
        part.ClearSelection2(True)
        part.WindowRedraw()

        View = Draw.GetFirstView 'A4P
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 104 / 1000 And Annpos(0) < 153 / 1000 And Annpos(1) > 43 / 1000 And Annpos(1) < 61 / 1000 Then
                        N1.SetText(material)
                    End If
                    If Annpos(0) > 153 / 1000 And Annpos(0) < 205 / 1000 And Annpos(1) > 43 / 1000 And Annpos(1) < 61 / 1000 Then
                        N1.SetText(company)
                    End If
                    If Annpos(0) > 153 / 1000 And Annpos(0) < 205 / 1000 And Annpos(1) > 23 / 1000 And Annpos(1) < 43 / 1000 Then
                        N1.SetText(name)
                    End If
                    If Annpos(0) > 153 / 1000 And Annpos(0) < 205 / 1000 And Annpos(1) > 12 / 1000 And Annpos(1) < 24 / 1000 Then
                        N1.SetText(Code)
                    End If
                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(Code)
                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop
        Draw.EditRebuild3()
    End Sub
    Public Sub A3模板(material$, company$, name$, Code$)
        Draw = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\gb_a3.drwdot", 1, 0, 0)
        part = swapp.ActiveDoc
        myNote = part.InsertNote("公司名称")
        If Not myNote Is Nothing Then
            myNote.LockPosition = False
            myNote.Angle = 0
            myNote.SetBalloon(0, 0)
            Annotation = myNote.GetAnnotation()
            If Not Annotation Is Nothing Then
                Annotation.SetLeader3(SwConst.swLeaderStyle_e.swNO_LEADER, 0, True, False, False, False)
                Annotation.SetPosition(0.375, 0.055, 0)
                Annotation.SetTextFormat(0, True, 0)
            End If
        End If
        part.ClearSelection2(True)
        part.WindowRedraw()

        View = Draw.GetFirstView 'A3
        Do Until View Is Nothing
            Notes = View.GetNotes()
            Count = View.GetNoteCount()
            If Count > 0 Then
                For Each N1 In Notes
                    Annotation = N1.GetAnnotation()
                    Annpos = Annotation.GetPosition()
                    If Annpos(0) > 315 / 1000 And Annpos(0) < 365 / 1000 And Annpos(1) > 43 / 1000 And Annpos(1) < 61 / 1000 Then
                        N1.SetText(material)
                    End If


                    If Annpos(0) > 365 / 1000 And Annpos(0) < 415 / 1000 And Annpos(1) > 42 / 1000 And Annpos(1) < 61 / 1000 Then
                        N1.SetText(company)
                    End If


                    If Annpos(0) > 365 / 1000 And Annpos(0) < 415 / 1000 And Annpos(1) > 23 / 1000 And Annpos(1) < 43 / 1000 Then
                        N1.SetText(name)
                    End If

                    If Annpos(0) > 365 / 1000 And Annpos(0) < 415 / 1000 And Annpos(1) > 14 / 1000 And Annpos(1) < 23 / 1000 Then
                        N1.SetText(Code)
                    End If


                    If Annpos(0) > 25 / 1000 And Annpos(0) < 85 / 1000 And Annpos(1) > 280 / 1000 And Annpos(1) < 292 / 1000 Then
                        N1.SetText(Code)
                    End If

                Next
            End If
            View = View.GetNextView() '获得下一个视图引用
        Loop
        Draw.EditRebuild3()
    End Sub
    Public Sub 设置图纸比例(图纸比例$)
        Dim Sheet1 As SldWorks.Sheet '图纸对象
        swapp = CreateObject("Sldworks.application")
        Draw = swapp.ActiveDoc
        Sheet1 = Draw.GetCurrentSheet()
        Dim SheetPr() As Double
        SheetPr = Sheet1.GetProperties2()

        Dim aaa As Object()
        aaa = 索引字符串(图纸比例$, ":")
        SheetPr(2) = aaa(0)
        SheetPr(3) = aaa(1) '分母
        SheetPr(4) = 1
        Sheet1.SetProperties2(SheetPr(0), SheetPr(1), SheetPr(2), SheetPr(3), SheetPr(4), SheetPr(5), SheetPr(6), SheetPr(7))
        Draw.EditRebuild()
    End Sub
    Public Function 索引字符串(str$, Optional 分隔符$ = "\") As Object()
        Dim strs() As Object
        索引字符串 = Split(str$, 分隔符$)
        '索引字符串 = 索引字符串
        '次序 = 次序 - 1
        '索引字符串 = strs(次序)
    End Function
    Public Sub 技术要求(X#, Y#, 技术要求1$, 技术要求2$, 技术要求3$, Optional 技术要求4$ = "", Optional 技术要求5$ = "", Optional 技术要求6$ = "", Optional 技术要求7$ = "", Optional 技术要求8$ = "", Optional 技术要求9$ = "", Optional 技术要求10$ = "")
        part = swapp.ActiveDoc
        part.FontPoints(18)
        myNote = part.InsertNote(技术要求1 + Chr(13) + Chr(10) +
           技术要求2 + Chr(13) + Chr(10) +
           技术要求3 + Chr(13) + Chr(10) +
           技术要求4 + Chr(13) + Chr(10) +
           技术要求5 + Chr(13) + Chr(10) +
           技术要求6 + Chr(13) + Chr(10) +
            技术要求7 + Chr(13) + Chr(10) +
            技术要求8 + Chr(13) + Chr(10) +
            技术要求9 + Chr(13) + Chr(10) +
            技术要求10)
        Annotation = myNote.GetAnnotation()
        Annotation.SetPosition(X, Y, 0)
    End Sub

    Public Sub 偏移注释位置(Annotation As SldWorks.Annotation, X偏移#, Y偏移#, Optional Z偏移# = 0)
        Dim dou As Double()
        dou = Annotation.GetPosition()
        Annotation.SetPosition2(dou(0) + X偏移, dou(1) + Y偏移, dou(2) + Z偏移)
    End Sub
    Public Sub 插入BOM表(view As SldWorks.View, Type%)
        '插入BOM表
        Dim swBOMTable As SldWorks.TableAnnotation
        Dim swBOMTable1 As SldWorks.TableAnnotation
        Dim width As Double
        Dim height As Double

        swBOMTable = view.InsertBomTable4(True, 0, 0, 1, Type, "默认", "D:\SOLIDWORKS\SOLIDWORKS\lang\chinese-simplified\bom-standard.sldbomtbt", False, 0, False) 'trpe=1只显示零件,2只显示顶层
        If swBOMTable Is Nothing Then
            swBOMTable = view.InsertBomTable4(True, 0, 0, 1, Type, "默认", "C:\PROGRA~1\SOLIDW~1\SOLIDW~1\lang\chinese-simplified\bom-standard.sldbomtbt", False, 0, False)
        End If

        swBOMTable.Anchored = False                      '获取或设置此表是否附加到其锚,如果表已附加到锚点,则为True
        swBOMTable.SetHeader(2, 4)                       '设置该表的标题,"2"底部,"4"表头中的行数
        swBOMTable.SetColumnType3(1, 204, False, "名称") '从0开始计算列数   '设置指定BOM表表列的类型和属性数据。"1"第一列,"204"定义的列类型:自定义属性,"false"将隐藏列不包含在索引中,"名称"数据属性
        swBOMTable.SetColumnType3(2, 204, False, "材料")
        swBOMTable.AnchorType = 4                        '获取或设置此表注释的锚点,"4"右侧底部
        swBOMTable.InsertColumn2(3, 0, "代号", 0)        '插入一列,名叫代号,"3"after,"0"位置,"0"定义列的宽度样式：默认宽度
        swBOMTable.SetColumnType3(1, 204, False, "代号")
        swBOMTable.InsertColumn2(3, 4, "单件", 2)          '"2"多线式
        swBOMTable.InsertColumn2(3, 5, "总计", 0)
        swBOMTable.InsertColumn2(3, 6, "备注", 0)
        width = swBOMTable.GetColumnWidth(0)
        swBOMTable.SetColumnWidth(0, width / 2, 0)
        width = swBOMTable.GetColumnWidth(1)
        swBOMTable.SetColumnWidth(1, width * 3 / 2, 1）
        width = swBOMTable.GetColumnWidth(2)
        swBOMTable.SetColumnWidth(2, width * 8 / 7, 0）
        width = swBOMTable.GetColumnWidth(3)
        swBOMTable.SetColumnWidth(3, width / 2, 0）
        width = swBOMTable.GetColumnWidth(4)
        swBOMTable.SetColumnWidth(4, width / 2, 0）
        width = swBOMTable.GetColumnWidth(7)
        swBOMTable.SetColumnWidth(7, width * 2, 0）


        Dim i As Integer
        i = 0
        Do Until i > swBOMTable.RowCount                 '获取此表的行数
            height = swBOMTable.GetRowHeight(i)          '获取指定行的高度
            swBOMTable.SetRowHeight(i, height * 0.8, 0)  '设置此表中指定行的高度
            i = i + 1
        Loop
        swBOMTable.Text2(i - 2, 0, False) = "序号"       '修改左下角表头/ '更改指定单元格的内容,行、列

        'Debug.Print(i)

        If i > 13 Then
            swBOMTable1 = swBOMTable.Split(2, i - 13) '行之后打断/将底层表赋给swBOMTable1,顶层表为swBOMTable

            Dim Sheet1 As SldWorks.Sheet '图纸对象
            Sheet1 = Draw.GetCurrentSheet()
            Dim BOMTableAnchor As SldWorks.TableAnchor
            part = Draw
            part.Extension.SelectByID2(Sheet1.GetName(), "SHEET", 0, 0, 0, False, 0, Nothing, 0)
            Draw.EditTemplate()                                                              '将当前工程图图纸的模板置于编辑模式
            'part.EditSketch()                                                               '允许在选定的工程图视图或图纸中编辑草图
            Dim skPoint As SldWorks.SketchPoint
            skPoint = part.SketchManager.CreatePoint(0.235, 0.005, 0)
            skPoint.Select4(False, Nothing)
            'oldTableAnchor = Sheet1.TableAnchor(2)                                          '获取指定的表锚,"2"材料清单
            BOMTableAnchor = Sheet1.SetAsTableAnchor(2)                                      '在工作表格式的选定点设置指定表的锚点,"2"材料清单
            swBOMTable.AnchorType = 4                                                        '获取或设置此表注释的锚点,"4"右侧底部
            swBOMTable.Anchored = True                                                    '获取或设置此表是否附加到其锚
            swBOMTable.Anchored = False


            '    Dim skPoint1 As SldWorks.SketchPoint
            '    skPoint1 = part.SketchManager.CreatePoint(0.415, 0.061, 0)
            '    skPoint1.Select4(False, Nothing)
            '    Sheet1.SetAsTableAnchor(2)                                      '在工作表格式的选定点设置指定表的锚点,"2"材料清单
            '    swBOMTable.AnchorType = 4                                                        '获取或设置此表注释的锚点,"4"右侧底部
            '    swBOMTable.Anchored = False                                                      '获取或设置此表是否附加到其锚

            Draw.EditSheet()
            'part.EditSketch()
            part.ForceRebuild3(True)
        End If









    End Sub

    Public Sub 插入零件标号(view As SldWorks.View, Optional 位置% = 3)
        '插入零件标号
        part.Extension.SelectByID2(view.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        Dim vNotes（） As Object
        Dim autoballoonParams As SldWorks.AutoBalloonOptions
        autoballoonParams = part.CreateAutoBalloonOptions()
        autoballoonParams.Layout = 位置
        autoballoonParams.ReverseDirection = False
        autoballoonParams.IgnoreMultiple = True
        autoballoonParams.InsertMagneticLine = False
        autoballoonParams.LeaderAttachmentToFaces = False
        autoballoonParams.Style = 10
        autoballoonParams.Size = 2
        autoballoonParams.EditBalloonOption = 1
        autoballoonParams.EditBalloons = 1
        autoballoonParams.UpperTextContent = 1
        autoballoonParams.UpperText = """"
        autoballoonParams.Layername = "尺寸"
        vNotes = part.AutoBalloon5(autoballoonParams)
        part.ClearSelection2（True）
    End Sub
    Public Function Initial_setting(type#)
        If type = 0 Then '关闭捕捉
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, False) '捕捉
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
        ElseIf type = 1 Then '激活捕捉,打开端点和草图点
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
        ElseIf type = 2 Then '激活捕捉,打开端点和草图点
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, True) '靠近
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
        ElseIf type = 3 Then '激活捕捉,打开端点和草图点
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInference, True) '捕捉
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, True)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPoints, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsMidPoints, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsIntersections, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsNearest, False) '靠近
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsTangent, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsParallel, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVLines, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsHVPoints, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsLength, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsGrid, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapsAngle, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchInferFromModel, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchAutomaticRelations, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swSketchSnapToGridIfDisplayed, False)
            swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '取消输入尺寸值
        End If

    End Function

End Class
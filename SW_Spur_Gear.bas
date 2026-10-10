Attribute VB_Name = "平歯車1"
' =====================================================
' SOLIDWORKS VBA: インボリュート平歯車の自動作成
' 動作確認: SOLIDWORKS 2026 (環境により要調整)
' 入力: 歯数 / モジュール(mm) / 歯幅(mm)
' 処理: 歯先円の押し出し → 歯溝カット → 中心軸作成 → 円周パターン
' 既知の制限:
'   ・歯底は基礎円より内側を直線で近似(トロコイドではない)
'   ・FeatureCircularPattern5 は API 上 Obsolete
'   ・基準平面名は 正面/Front Plane, 右側面/Right Plane, 平面/Top Plane を想定
' =====================================================
Option Explicit

Const PI As Double = 3.14159265358979
Const PRESSURE_ANGLE As Double = 0.349065850398866   ' 圧力角 20° (rad)
Const SPLINE_PTS As Long = 15                        ' インボリュートの分割点数

' 歯車パラメータ(関数から参照するためモジュールレベルで保持)
Dim gN As Long
Dim gM As Double
Dim gRp As Double
Dim gRa As Double
Dim gRb As Double
Dim gRf As Double

' =====================================================
' メイン処理
' =====================================================
Sub main()
    ' --- 【パラメータ入力】 ---
    Dim inputText As String
    Dim modulus As Double
    Dim thickness As Double

    ' 歯数
    inputText = InputBox("歯数を入力してください (6?200)", "歯数の入力", "20")
    If Len(inputText) = 0 Then Exit Sub        ' キャンセル時は終了
    If Not IsNumeric(inputText) Then
        MsgBox "歯数は数値で入力してください。"
        Exit Sub
    End If
    gN = CLng(inputText)
    If gN < 6 Or gN > 200 Then
        MsgBox "歯数は 6?200 の範囲で入力してください。"
        Exit Sub
    End If

    ' モジュール (mm)
    inputText = InputBox("モジュール(mm)を入力してください", "モジュールの入力", "2")
    If Len(inputText) = 0 Then Exit Sub
    If Not IsNumeric(inputText) Then
        MsgBox "モジュールは数値で入力してください。"
        Exit Sub
    End If
    modulus = Val(inputText)
    If modulus <= 0# Then
        MsgBox "モジュールは 0 より大きい値を入力してください。"
        Exit Sub
    End If

    ' 歯幅 (mm)
    inputText = InputBox("歯幅(mm)を入力してください", "歯幅の入力", "15")
    If Len(inputText) = 0 Then Exit Sub
    If Not IsNumeric(inputText) Then
        MsgBox "歯幅は数値で入力してください。"
        Exit Sub
    End If
    thickness = Val(inputText)
    If thickness <= 0# Then
        MsgBox "歯幅は 0 より大きい値を入力してください。"
        Exit Sub
    End If
    ' --------------------------
    ' --------------------------

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "SOLIDWORKSに接続できませんでした。"
        Exit Sub
    End If

    ' ---- 0. 新規部品を作成 ----
    Dim defaultTemplate As String
    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)

    Dim swModel As SldWorks.ModelDoc2
    If Len(defaultTemplate) > 0 Then
        Set swModel = swApp.NewDocument(defaultTemplate, 0, 0, 0)
    End If
    If swModel Is Nothing Then
        Set swModel = swApp.NewPart
    End If
    If swModel Is Nothing Then
        MsgBox "新規部品を開けませんでした。" & vbCrLf & _
               "テンプレートパス: [" & defaultTemplate & "]"
        Exit Sub
    End If

    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim swSketchMgr As SldWorks.SketchManager
    Dim swFeatureMgr As SldWorks.FeatureManager
    Set swModelDocExt = swModel.Extension
    Set swSketchMgr = swModel.SketchManager
    Set swFeatureMgr = swModel.FeatureManager

    ' 寸法計算 (SOLIDWORKSの内部単位はメートル)
    gM = modulus / 1000#
    Dim t As Double: t = thickness / 1000#
    gRp = gM * gN / 2#                 ' ピッチ円半径
    gRa = gRp + gM                     ' 歯先円半径
    gRf = gRp - 1.25 * gM              ' 歯底円半径
    gRb = gRp * Cos(PRESSURE_ANGLE)    ' 基礎円半径

    ' ---- 1. 円盤スケッチ(歯先円) ----
    If Not SelectFirstRefPlane(swModel) Then
        MsgBox "基準平面が見つかりませんでした。" & vbCrLf & vbCrLf & _
               "ツリーにある要素: " & vbCrLf & ListFeatureNames(swModel)
        Exit Sub
    End If
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCircle 0, 0, 0, 0, gRa, 0
    swSketchMgr.InsertSketch True

    ' ---- 2. 円盤を押し出し ----
    swModel.ClearSelection2 True
    SelectLastSketch swModel
    swFeatureMgr.FeatureExtrusion3 True, False, False, 0, 0, t, 0, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False

    ' ---- 3. 歯溝スケッチ ----
    swModel.ClearSelection2 True
    If Not SelectFirstRefPlane(swModel) Then
        MsgBox "基準平面が見つかりませんでした(歯溝スケッチ)。"
        Exit Sub
    End If
    swSketchMgr.InsertSketch True
    DrawToothGap swSketchMgr
    swSketchMgr.InsertSketch True

    ' ---- 4. 歯溝をカット(カット方向 Dir を切り替えて再試行) ----
    Dim swCutFeat As SldWorks.Feature
    Dim dirs As Variant
    Dim k As Long
    dirs = Array(True, False)   ' 既定の反対方向 → 既定方向 の順に試す
    For k = LBound(dirs) To UBound(dirs)
        swModel.ClearSelection2 True
        SelectLastSketch swModel
        Set swCutFeat = swFeatureMgr.FeatureCut4(True, False, dirs(k), 0, 0, t, 0, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
        If Not swCutFeat Is Nothing Then Exit For
    Next k
    If swCutFeat Is Nothing Then
        MsgBox "歯溝のカットに失敗しました。" & vbCrLf & _
               "・スケッチが選択されているか" & vbCrLf & _
               "・円盤の押し出し方向とスケッチの位置関係" & vbCrLf & _
               "を確認してください。"
        Exit Sub
    End If

    ' ---- 5. 中心軸を作成して円周パターン ----
    swModel.ClearSelection2 True
    If Not CreateCenterAxis(swModel) Then
        MsgBox "中心軸を作成できませんでした。"
        Exit Sub
    End If

    If Not SelectLastAxis(swModel) Then
        MsgBox "作成した中心軸を選択できませんでした。"
        Exit Sub
    End If
    swCutFeat.Select2 True, 4
    swFeatureMgr.FeatureCircularPattern5 gN, 2# * PI, True, "NULL", False, True, False, False, False, False, 0, 0, "NULL", False

    swModel.ViewZoomtofit2
    MsgBox "歯車の作成が完了しました。"
End Sub

' =====================================================
' 歯溝の輪郭(閉じたループ)を描画
' =====================================================
Private Sub DrawToothGap(ByVal swSketchMgr As SldWorks.SketchManager)
    Dim rO As Double: rO = gRa + 0.5 * gM           ' 歯先より外側まで伸ばす
    Dim thA As Double: thA = FlankAngle(gRb, 0)     ' A: 右フランク 基礎円
    Dim thB As Double: thB = FlankAngle(rO, 0)      ' B: 右フランク 外側
    Dim phC As Double: phC = FlankAngle(rO, 1)      ' C: 左フランク 外側
    Dim phD As Double: phD = FlankAngle(gRb, 1)     ' D: 左フランク 基礎円

    Dim pts As Variant

    ' 右フランク A → B
    pts = BuildFlankPoints(gRb, rO, 0)
    swSketchMgr.CreateSpline pts

    ' 外側円弧 B → C
    swSketchMgr.CreateArc 0, 0, 0, rO * Cos(thB), rO * Sin(thB), 0, rO * Cos(phC), rO * Sin(phC), 0, 1

    ' 左フランク C → D
    pts = BuildFlankPoints(rO, gRb, 1)
    swSketchMgr.CreateSpline pts

    ' 径方向線 D → E (歯底円まで)
    swSketchMgr.CreateLine gRb * Cos(phD), gRb * Sin(phD), 0, gRf * Cos(phD), gRf * Sin(phD), 0

    ' 歯底円弧 E → F
    swSketchMgr.CreateArc 0, 0, 0, gRf * Cos(phD), gRf * Sin(phD), 0, gRf * Cos(thA), gRf * Sin(thA), 0, -1

    ' 径方向線 F → A
    swSketchMgr.CreateLine gRf * Cos(thA), gRf * Sin(thA), 0, gRb * Cos(thA), gRb * Sin(thA), 0
End Sub

' =====================================================
' インボリュートの点列(x,y,z の一次元配列)を生成
' =====================================================
Private Function BuildFlankPoints(ByVal r1 As Double, ByVal r2 As Double, ByVal side As Long) As Double()
    Dim arr() As Double
    ReDim arr(3 * SPLINE_PTS - 1)
    Dim i As Long, r As Double, a As Double
    For i = 0 To SPLINE_PTS - 1
        If i = 0 Then
            r = r1
        ElseIf i = SPLINE_PTS - 1 Then
            r = r2
        Else
            r = r1 + (r2 - r1) * i / (SPLINE_PTS - 1)
        End If
        a = FlankAngle(r, side)
        arr(3 * i) = r * Cos(a)
        arr(3 * i + 1) = r * Sin(a)
        arr(3 * i + 2) = 0#
    Next i
    BuildFlankPoints = arr
End Function

' =====================================================
' 半径 r におけるフランクの極角 (side 0: 右, 1: 左)
' =====================================================
Private Function FlankAngle(ByVal r As Double, ByVal side As Long) As Double
    Dim th As Double
    th = PI / (2# * gN) - (InvAng(Acos(gRb / r)) - InvAng(PRESSURE_ANGLE))
    If side = 0 Then
        FlankAngle = th
    Else
        FlankAngle = 2# * PI / gN - th
    End If
End Function

' インボリュート関数 inv(a) = tan(a) - a  (VBAにTanはないので sin/cos で計算)
Private Function InvAng(ByVal a As Double) As Double
    InvAng = Sin(a) / Cos(a) - a
End Function

' Acos (VBAに標準関数がないため定義)
Private Function Acos(ByVal x As Double) As Double
    If x >= 1# Then
        Acos = 0#
    Else
        Acos = Atn(-x / Sqr(-x * x + 1#)) + 2# * Atn(1#)
    End If
End Function

' =====================================================
' 基準平面(正面)を名前で選択
' =====================================================
Private Function SelectFirstRefPlane(ByVal model As SldWorks.ModelDoc2) As Boolean
    Dim names As Variant
    Dim i As Long
    Dim ok As Boolean

    names = Array("正面", "Front Plane")
    ok = False

    If Not model Is Nothing Then
        For i = LBound(names) To UBound(names)
            model.ClearSelection2 True
            If model.Extension.SelectByID2(names(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0) Then
                ok = True
                Exit For
            End If
        Next i
    End If

    SelectFirstRefPlane = ok
End Function

' =====================================================
' 右側面と平面(上面)の交線として中心軸を作成
' =====================================================
Private Function CreateCenterAxis(ByVal model As SldWorks.ModelDoc2) As Boolean
    Dim rightNames As Variant, topNames As Variant
    Dim rightOk As Boolean, topOk As Boolean
    Dim i As Long
    Dim countBefore As Long, countAfter As Long

    CreateCenterAxis = False
    If model Is Nothing Then Exit Function

    rightNames = Array("右側面", "Right Plane")
    topNames = Array("平面", "Top Plane")

    ' 1枚目: 右側面(Append=False で選択をリセット)
    model.ClearSelection2 True
    For i = LBound(rightNames) To UBound(rightNames)
        If model.Extension.SelectByID2(rightNames(i), "PLANE", 0, 0, 0, False, 1, Nothing, 0) Then
            rightOk = True
            Exit For
        End If
    Next i

    ' 2枚目: 上面(平面)を追加選択(Append=True)
    For i = LBound(topNames) To UBound(topNames)
        If model.Extension.SelectByID2(topNames(i), "PLANE", 0, 0, 0, True, 1, Nothing, 0) Then
            topOk = True
            Exit For
        End If
    Next i

    If Not (rightOk And topOk) Then
        MsgBox "平面を選択できませんでした。" & vbCrLf & _
               "右側面: " & rightOk & " / 上面: " & topOk & vbCrLf & vbCrLf & _
               "ツリーにある要素: " & vbCrLf & ListFeatureNames(model)
        model.ClearSelection2 True
        Exit Function
    End If

    ' 作成前の軸の数を数える
    countBefore = CountAxes(model)

    ' 基準軸を挿入。戻り値は信用せず、後で軸の数で判定する
    model.InsertAxis2 True

    model.ClearSelection2 True

    ' 軸の数が増えたかで成否を判定
    countAfter = CountAxes(model)
    If countAfter > countBefore Then
        CreateCenterAxis = True
    Else
        MsgBox "中心軸が作成されませんでした。" & vbCrLf & _
               "作成前の軸数: " & countBefore & " / 作成後: " & countAfter
    End If
End Function

' =====================================================
' 基準軸の数を数える
' =====================================================
Private Function CountAxes(ByVal model As SldWorks.ModelDoc2) As Long
    Dim f As SldWorks.Feature
    Dim n As Long
    If model Is Nothing Then Exit Function
    Set f = model.FirstFeature
    Do While Not f Is Nothing
        If IsAxisFeature(f) Then n = n + 1
        Set f = f.GetNextFeature
    Loop
    CountAxes = n
End Function

' =====================================================
' 基準軸のフィーチャーか判定(型名 または 名前で判定)
' =====================================================
Private Function IsAxisFeature(ByVal f As SldWorks.Feature) As Boolean
    Dim tn As String, nm As String
    tn = f.GetTypeName2
    nm = f.Name
    IsAxisFeature = (tn = "RefAxis") Or (Left(nm, 1) = "軸") Or _
                    (Left(nm, 3) = "基準軸") Or (Left(nm, 4) = "Axis")
End Function

' =====================================================
' 作成した中心軸を選択(選択数で成否を判定)
' =====================================================
Private Function SelectLastAxis(ByVal model As SldWorks.ModelDoc2) As Boolean
    Dim f As SldWorks.Feature, lastF As SldWorks.Feature
    Dim lastName As String
    Dim r1 As Boolean, r2 As Boolean
    Dim cnt As Long

    SelectLastAxis = False
    If model Is Nothing Then Exit Function

    Set f = model.FirstFeature
    Do While Not f Is Nothing
        If IsAxisFeature(f) Then
            Set lastF = f
            lastName = f.Name
        End If
        Set f = f.GetNextFeature
    Loop
    If lastF Is Nothing Then
        MsgBox "軸のフィーチャーが見つかりませんでした。"
        Exit Function
    End If

    ' 方法1: 名前で選択
    model.ClearSelection2 True
    r1 = model.Extension.SelectByID2(lastName, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
    cnt = model.SelectionManager.GetSelectedObjectCount2(-1)

    ' 方法2: フィーチャーから選択(方法1で選択数が1にならなかった場合)
    If cnt <> 1 Then
        model.ClearSelection2 True
        r2 = lastF.Select2(False, 1)
        cnt = model.SelectionManager.GetSelectedObjectCount2(-1)
    End If

    ' 選択数が1なら成功とみなす
    If cnt = 1 Then
        SelectLastAxis = True
    Else
        MsgBox "軸の選択に失敗しました。" & vbCrLf & _
               "対象: " & lastName & vbCrLf & _
               "名前での選択: " & r1 & " / フィーチャーでの選択: " & r2 & vbCrLf & _
               "選択数: " & cnt
    End If
End Function

' =====================================================
' 最後に作成されたスケッチを選択
' =====================================================
Private Sub SelectLastSketch(ByVal model As SldWorks.ModelDoc2)
    If model Is Nothing Then Exit Sub
    Dim f As SldWorks.Feature, lastF As SldWorks.Feature
    Set f = model.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "ProfileFeature" Then Set lastF = f
        Set f = f.GetNextFeature
    Loop
    If Not lastF Is Nothing Then lastF.Select2 False, 0
End Sub

' =====================================================
' デバッグ用: ツリーのフィーチャー名を一覧にする
' =====================================================
Private Function ListFeatureNames(ByVal model As SldWorks.ModelDoc2) As String
    Dim f As SldWorks.Feature
    Dim s As String
    If model Is Nothing Then
        ListFeatureNames = "(モデルなし)"
        Exit Function
    End If
    Set f = model.FirstFeature
    Do While Not f Is Nothing
        s = s & f.Name & " (" & f.GetTypeName2 & ")" & vbCrLf
        Set f = f.GetNextFeature
    Loop
    ListFeatureNames = s
End Function


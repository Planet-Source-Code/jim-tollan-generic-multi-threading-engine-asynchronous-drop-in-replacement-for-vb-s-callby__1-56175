Attribute VB_Name = "mdlIDispatch"
' ************************************************************************
' Copyright:    All rights reserved.  © 2004
' Project:      AsyncServer
' Module:       mdlIDispatch
' Author:       james b tollan
' Date:         15-Sep-2004 14:34
' Revisions **************************************************************
' Date          Name                Description
' ************************************************************************

Option Explicit

Public Function CallByNameFixParamArray _
    (pobjTarget As Object, _
    ByVal pstrProcName As Variant, _
    ByVal CallType As cbnCallTypes, _
    ParamArray pArgs()) As Variant

    ' pobjTarget    :   Class or form object that contains the procedure/property
    ' pstrProcName  :   Name of the procedure or property
    ' CallType      :   vbLet/vbGet/vbSet/vbMethod
    ' pArgs()       :   Param Array of parameters required for methode/property
    ' usage             CallByNameFixParamArray(CTest, "AddNumbers", vbMethod, 1, 2, 3, CAnotherClass)

    Dim IDsp        As IDispatch.IDispatchVB
    Dim rIid        As IDispatch.iid
    Dim Params      As IDispatch.DISPPARAMS
    Dim Excep       As IDispatch.EXCEPINFO
    ' Do not remove TLB because those types
    ' are also defined in stdole
    Dim DISPID      As Long
    Dim lngArgErr   As Long
    Dim varRet      As Variant
    Dim varArr()    As Variant

    Dim lngRet      As Long
    Dim lngLoop     As Long
    Dim lngMax      As Long

    ' Get IDispatch from object
    Set IDsp = pobjTarget

    ' Get DISPIP from pstrProcName
    lngRet = IDsp.GetIDsOfNames(rIid, StrConv(pstrProcName, vbUnicode), 1&, 0&, DISPID)

    If lngRet = 0 Then

        If Not IsMissing(pArgs) Then
            lngMax = UBound(pArgs(0))
            If lngMax >= 0 Then
                ReDim varArr(0 To lngMax)
                ' Fill parameters arrays. The array must be
                ' filled in reverse order.
                For lngLoop = 0 To lngMax
                    VariantCopy varArr(lngMax - lngLoop), pArgs(0)(lngLoop)
                Next
                With Params
                    .cArgs = UBound(varArr) + 1
                    .rgPointerToVariantArray = VarPtr(varArr(0))
                End With
            End If
        End If

        ' Invoke method/property
        lngRet = IDsp.Invoke(DISPID, rIid, 0, CallType, Params, varRet, Excep, lngArgErr)

        If lngRet <> 0 Then
            If lngRet = DISP_E_EXCEPTION Then
                Err.Raise Excep.wCode
            Else
                Err.Raise lngRet
            End If
        End If
    Else
        Err.Raise lngRet
    End If

    On Error Resume Next

    Set IDsp = Nothing
    Set CallByNameFixParamArray = varRet

    If Err.Number <> 0 Then CallByNameFixParamArray = varRet

End Function

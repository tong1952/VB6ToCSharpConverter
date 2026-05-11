VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private WithEvents mobjCommands As CommandListLib.Commands
Attribute mobjCommands.VB_VarHelpID = -1
Private WithEvents mobjStatuses As TalentProcessLib.Statuses
Attribute mobjStatuses.VB_VarHelpID = -1
Private mobjConstants As ControlLib.ControlConstants
Private mobjGlobalVars As ControlLib.GlobalVars
Private mobjStats As AdvUtilLib.Stats
Private WithEvents tmrMain As LoopTimer
Attribute tmrMain.VB_VarHelpID = -1
Private WithEvents tmrRunOnce As ccrpTimer
Attribute tmrRunOnce.VB_VarHelpID = -1
Private tmrElapsed As ccrpStopWatch
Private WithEvents mobjInterlocks As Interlocks
Attribute mobjInterlocks.VB_VarHelpID = -1
Private mobjEvents As TalentProcessLib.Events
Private mobjForm As frmControl
Private mobjSPGConstants As SharedPropertyGroup
Private mobjSPGDataItems As SharedPropertyGroup
Private mobjParManager As ParametersLib.ParManager

Implements ICcrpTimerNotify

Private mlID As Long
Private mbReady As Boolean
Private mbInitialized As Boolean
Private Const bKeepStats As Boolean = False
Private bRealTime As Boolean
Private mvInterlockCheckValue As Variant
Private mdStartTime As Date
Private Const mbKeepStats As Boolean = True
Private mbResetSim As Boolean

Private mfMFSpeedSPBuffer() As Single   ' fan delay pipeline
Private mlMFSpeedSPBufferIndex As Long
Private mlMFSpeedSPBufferSize As Long

Private mlNumFilter1 As Long
Private mlNumRampTowards As Long
Private mlNumFollow As Long
Private mlNumFollow2 As Long

' Invalid Opcode Collections
Private mcolInvalidOpcodes As Collection

Private Enum InitTypes
    initAll = 0
    initConstants = 1
    initPars = 2
End Enum

'' Module level control constants
Private mccsPriority As String
Private mccfUpdateTime As Single

' >>> Beginning of [Control Constants] declaration - do not move or change this line
Private mccfFanDelay As Single
Private mccfTunnelVolume As Single
' <<< End of [Control Constants] declaration - do not move or change this line

' Other control constants
Private mccfKp1Coeff(0 To 2) As Single
Private mccfKp2Coeff(0 To 2) As Single
Private mccfKpDpTrans As Single
Private mccfKq1Coeff(0 To 2) As Single
Private mccfKq2Coeff(0 To 2) As Single
Private mccfKqDpTrans As Single
Private mccfDPxOverHi As Single
Private mccfDPxOverLo As Single

' Module level Inputs
' >>> Beginning of [Input Parameters] declaration - do not move or change this line
Private mipAirDensity As Single
Private mipAT_Rate As Single
Private mipCCE_Start As Single
Private mipCS_BrineReq As Single
Private mipCS_HtrSetpt As Single
Private mipCV_Setpt As Single
Private mipDCA_Start As Single
Private mipDCA_Target As Single
Private mipDewpoint As Single
Private mipDP_Target As Single
Private mipDV_Gly_Start As Single
Private mipDV_Gly_TTgt As Single
Private mipDV_Oil_Start As Single
Private mipDV_Oil_TTgt As Single
Private mipDy_CommSts As Single
Private mipDy_Remote As Single
Private mipDyPit_Start As Single
Private mipDyPit_Target As Single
Private mipEX_FanSetpt As Single
Private mipEX_Start As Single
Private mipHR_BrnSetpt As Single
Private mipHR_RunSts As Single
Private mipHR1_CVSetpt As Single
Private mipHR1_HtrSetpt As Single
Private mipHR1_RunSts As Single
Private mipHR2_CVSetpt As Single
Private mipHR2_HtrSetpt As Single
Private mipHR2_RunSts As Single
Private mipHR3_CVSetpt As Single
Private mipHR3_HtrSetpt As Single
Private mipHR3_RunSts As Single
Private mipHR4_CVSetpt As Single
Private mipHR4_HtrSetpt As Single
Private mipHR4_RunSts As Single
Private mipIDL_Start As Single
Private mipLTBrn_Warn As Single
Private mipMF_Setpt As Single
Private mipMF_Start As Single
Private mipMTBrn_Warn As Single
Private mipP201_RunReq As Single
Private mipPG_HL_DelRst As Single
Private mipPG_HL_Start As Single
Private mipPG_LL_Start As Single
Private mipPGDy_HL_DRst As Single
Private mipPGDy_HL_Strt As Single
Private mipPGDy_LL_Strt As Single
Private mipPStatic As Single
Private mipQ As Single
Private mipRV_Setpt As Single
Private mipSR1_RunReq As Single
Private mipSR1_Setpt As Single
Private mipSR2_RunReq As Single
Private mipSR2_Setpt As Single
Private mipSR3_RunReq As Single
Private mipSR3_Setpt As Single
Private mipSR4_RunReq As Single
Private mipSR4_Setpt As Single
Private mipSR5_RunReq As Single
Private mipSR5_Setpt As Single
Private mipSR6_RunReq As Single
Private mipSR6_Setpt As Single
Private mipSS_ArmExtReq As Single
Private mipSS_FanStart As Single
Private mipSS_InvCoef0 As Single
Private mipSS_InvCoef1 As Single
Private mipSS_InvCoef2 As Single
Private mipSS_InvCoef3 As Single
Private mipSS_InvCoef4 As Single
Private mipSS_Power As Single
Private mipSS_RunReq As Single
Private mipSV_Setpt As Single
Private mipVH_IgnEnable As Single
Private mipWS_Setpt As Single
' <<< End of [Input Parameters] declaration - do not move or change this line

'Private Const MAX_RAIN = 30
'Private mipRainValve(1 To MAX_RAIN) As Single

' Module level Outputs
' >>> Beginning of [Output Parameters] declaration - do not move or change this line
Private mopAirtemp As Single
Private mopAL1_FaultSts As Single
Private mopAL2_FaultSts As Single
Private mopBP_Raw As Single
Private mopCCE_Perm As Single
Private mopCCE_Remote As Single
Private mopCCE_RunSts As Single
Private mopCS_BrineSts As Single
Private mopCS_ChSumFlt As Single
Private mopCS_CritFault As Single
Private mopCS_GenFault As Single
Private mopCS_Health As Single
Private mopCS_HtrFltSts As Single
Private mopCS_HtrRunSts As Single
Private mopCS_LTclsSts As Single
Private mopCS_LTopnSts As Single
Private mopCS_MTclsSts As Single
Private mopCS_MTopnSts As Single
Private mopCS_RunSts As Single
Private mopCtrl_Power As Single
Private mopCV_Position As Single
Private mopDCA_FaultSts As Single
Private mopDCA_Perm As Single
Private mopDCA_Remote As Single
Private mopDCA_RunSts As Single
Private mopDCA_Temp As Single
Private mopDeltaP_Hi As Single
Private mopDeltaP_Lo As Single
Private mopDoor_Dy As Single
Private mopDoor_UpLg_AS As Single
Private mopDV_Gly_Temp As Single
Private mopDV_Oil_Temp As Single
Private mopDy_Estop_Sts As Single
Private mopDy_Fault As Single
Private mopDyPit_Temp As Single
Private mopEPB_Ctrl_Rm As Single
Private mopEPB_TS_NW As Single
Private mopEPB_TS_SW As Single
Private mopEPB_UpLeg_AS As Single
Private mopEStop_SumSts As Single
Private mopEX_FanSpeed As Single
Private mopEX_Fault As Single
Private mopEX_Perm As Single
Private mopEX_Remote As Single
Private mopEX_RunSts As Single
Private mopFaultBus As Single
Private mopFire_DYSup As Single
Private mopFire_Sum As Single
Private mopGas_Sys_Sts As Single
Private mopHD_Mirror As Single
Private mopHD_SmplFlow As Single
Private mopHR_Brn_Temp As Single
Private mopHR_EstopSts As Single
Private mopHR_FaultSts As Single
Private mopHR_Perm As Single
Private mopHR1_FaultSts As Single
Private mopHR1_HxInTmp As Single
Private mopHR1_HxOutTmp As Single
Private mopHR1_PnlTemp As Single
Private mopHR2_FaultSts As Single
Private mopHR2_HxInTmp As Single
Private mopHR2_HxOutTmp As Single
Private mopHR2_PnlTemp As Single
Private mopHR3_FaultSts As Single
Private mopHR3_HxInTmp As Single
Private mopHR3_HxOutTmp As Single
Private mopHR3_PnlTemp As Single
Private mopHR4_FaultSts As Single
Private mopHR4_HxInTmp As Single
Private mopHR4_HxOutTmp As Single
Private mopHR4_PnlTemp As Single
Private mopHX_HiDeltaP As Single
Private mopHxIn_Temp As Single
Private mopHxOut_Temp As Single
Private mopIDL_Fault As Single
Private mopIDL_FlCldSts As Single
Private mopIDL_FlOpdSts As Single
Private mopIDL_ModeOff As Single
Private mopIDL_ModeOn As Single
Private mopIDL_Perm As Single
Private mopIDL_Remote As Single
Private mopIDL_ShCldSts As Single
Private mopIDL_ShOpdSts As Single
Private mopKey_Dy_Door As Single
Private mopKey_Veh_Door As Single
Private mopLo_Plen_P As Single
Private mopLoLo_Plen_P As Single
Private mopLTBrn_Temp As Single
Private mopMF_Fault As Single
Private mopMF_Perm As Single
Private mopMF_Remote As Single
Private mopMF_RunSts As Single
Private mopMF_SpdLim1 As Single
Private mopMF_SpdLim2 As Single
Private mopMF_Speed As Single
Private mopMF_SupplySts As Single
Private mopMF_Warning As Single
Private mopMixBrn_Temp As Single
Private mopMTBrn_Temp As Single
Private mopMUA_DP As Single
Private mopMUA_Flow As Single
Private mopMUA_HX109Flt As Single
Private mopMUA_Temp As Single
Private mopP201_FltSts As Single
Private mopP201_RemSts As Single
Private mopP201_RunSts As Single
Private mopPabs As Single
Private mopPG_FanFault As Single
Private mopPG_FanPerm As Single
Private mopPG_FanRemote As Single
Private mopPG_FanRunH As Single
Private mopPG_FanRunL As Single
Private mopPG_HL_Tmr As Single
Private mopPG_HLActive As Single
Private mopPG_HLRunSts As Single
Private mopPGDy_FanFlt As Single
Private mopPGDy_FanPerm As Single
Private mopPGDy_FRunH As Single
Private mopPGDy_FRunL As Single
Private mopPGDy_HL_Tmr As Single
Private mopPGDy_HLRnSts As Single
Private mopPGDy_Remote As Single
Private mopRelHumidity As Single
Private mopRV_Position As Single
Private mopSR1_AirTemp As Single
Private mopSR1_FaultSts As Single
Private mopSR1_Master As Single
Private mopSR1_RunSts As Single
Private mopSR1_SetptFbk As Single
Private mopSR2_AirTemp As Single
Private mopSR2_FaultSts As Single
Private mopSR2_Master As Single
Private mopSR2_RunSts As Single
Private mopSR2_SetptFbk As Single
Private mopSR3_AirTemp As Single
Private mopSR3_FaultSts As Single
Private mopSR3_Master As Single
Private mopSR3_RunSts As Single
Private mopSR3_SetptFbk As Single
Private mopSR4_AirTemp As Single
Private mopSR4_FaultSts As Single
Private mopSR4_Master As Single
Private mopSR4_RunSts As Single
Private mopSR4_SetptFbk As Single
Private mopSR5_AirTemp As Single
Private mopSR5_FaultSts As Single
Private mopSR5_Master As Single
Private mopSR5_RunSts As Single
Private mopSR5_SetptFbk As Single
Private mopSR6_AirTemp As Single
Private mopSR6_FaultSts As Single
Private mopSR6_Master As Single
Private mopSR6_RunSts As Single
Private mopSR6_SetptFbk As Single
Private mopSS_ArmExtSts As Single
Private mopSS_ArmRetSts As Single
Private mopSS_Fan1Fault As Single
Private mopSS_Fan1Rem As Single
Private mopSS_Fan2Fault As Single
Private mopSS_Fan2Rem As Single
Private mopSS_FanPerm As Single
Private mopSS_Fn1RunSts As Single
Private mopSS_Fn2RunSts As Single
Private mopSS_HiTemp As Single
Private mopSS_InputPyro As Single
Private mopSS_Perm As Single
Private mopSS_RunSts As Single
Private mopSteamPress As Single
Private mopSV_Position As Single
Private mopTA1_FaultSts As Single
Private mopTA2_FaultSts As Single
Private mopVH_IgnPerm As Single
Private mopVH_IgnSts As Single
' <<< End of [Output Parameters] declaration - do not move or change this line

' Module level Global variables
Private mlDefPriority As Long
Private mgvfListParsMissing As Single

' >>> Beginning of [Global Variables] declaration - do not move or change this line
Private mgvfAT_Previous As Single
Private mgvfAT_Sim As Single
Private mgvfAT_Sim_002 As Single
Private mgvfAT_Sim_003 As Single
Private mgvfAT_Sim_004 As Single
Private mgvfAT_Sim_005 As Single
Private mgvfAT_Sim_006 As Single
Private mgvfCondensTemp As Single
Private mgvfDP_Sim As Single
Private mgvfDP_Sim_001 As Single
Private mgvfDP_Sim_002 As Single
Private mgvfDV_Sim As Single
Private mgvfDY_Sim As Single
Private mgvfEX_Sim As Single
Private mgvfFaultsSim As Single
Private mgvfHD_Sim As Single
Private mgvfHD_Sim_001 As Single
Private mgvfHD_Sim_002 As Single
Private mgvfHD_Sim_003 As Single
Private mgvfHR_Sim As Single
Private mgvfHx_AvgTemp As Single
Private mgvfIdle_Sim As Single
Private mgvfMF_Previous As Single
Private mgvfMF_Sim_001 As Single
Private mgvfMF_Sim_002 As Single
Private mgvfMF_SpeedDelayed As Single
Private mgvfMF_SpeedDelayType As Single
Private mgvfPG_Sim As Single
Private mgvfPstatic As Single
Private mgvfQ_Sim As Single
Private mgvfQ_Sim_001 As Single
Private mgvfResetSim As Single
Private mgvfSF_Sim As Single
Private mgvfSR_Sim As Single
Private mgvfSS_Sim As Single
Private mgvfTimeInterval_Sim As Single
Private mgvfTimeInterval_Sim_001 As Single
Private mgvfWS_Sim As Single
Private mgvfWS_Sim_001 As Single
' <<< End of [Global Variables] declaration - do not move or change this line

' Function Objects
Private mobjFilter1() As Filter1
Private mobjRampTowards() As Ramp1
Private mobjFollow() As Follow
Private mobjFollow2() As Follow2

Private Type SSLamp
    CloudShutterCmd As Single
    CloudShutterFault As Single
    CloudShutterStatus As Single
    Control As Single
    DimmingTime As Single
    FAULT As Single
    LampFault As Single
    Offset As Single
    OffsetSp As Single
    OperatingHours As Single
    OvertempError As Single
    PhaseError As Single
    Power As Single
    Ready As Single
    RestartTime As Single
    Selected As Single
    SelectedStatus As Single
    ShortCircuitError As Single
    Status As Single
    TunnelShutterCmd As Single
    TunnelShutterFault As Single
    TunnelShutterStatus As Single
End Type

Private Enum SolarModes
    smOff = 0
    smManual = 1        ' Power (%) control
    smAutoPLC = 2       ' Intensity control via PLC
    smAutoProfile = 3   ' Intensity control via scaling a loaded profile
End Enum

Private mudtSSLamp(1 To 28) As SSLamp

Private Sub Calc_AT_Sim()
    
    Const PROC_NAME = "Calc_AT_Sim"
    Dim fTavg As Single
    Dim fNozzArea As Single
    Dim fdTA As Single
    Dim fdTB As Single
    Dim fCapacity As Single
    Dim fTDelta As Single
    
    On Error GoTo Error
    
    mgvfAT_Previous = mobjRampTowards(0)(mopAirtemp, mgvfAT_Previous, mipAT_Rate / 60, mccfUpdateTime, 0.01)
    
    If mgvfAT_Sim = 0 Or mbResetSim Then
        mopCS_ChSumFlt = 1
        mopCS_CritFault = 1
        mopCS_GenFault = 1
        mopCS_Health = 1
        mopCS_HtrFltSts = 1
        mopCS_RunSts = 1
    
        mopLTBrn_Temp = -40
        mopMTBrn_Temp = -25
    
        mopHxIn_Temp = 20
        mopHxOut_Temp = 20
        mopMixBrn_Temp = mopMTBrn_Temp
        mopAirtemp = 20
        mopP201_RemSts = 1
        mopP201_FltSts = 1
    
        mgvfAT_Sim_002 = 0.0001
        mgvfAT_Sim_003 = 0.1
        mgvfAT_Sim_004 = 0.00001
        mgvfAT_Sim_005 = 0.001
        mgvfAT_Sim_006 = 0.0001
        
        mgvfAT_Sim = 1
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    If mopCS_BrineSts = 1 Then
        ' MT
        mopCS_LTopnSts = 0
        mopCS_LTclsSts = mobjFollow2(0)(1, mopCS_LTclsSts, 3)
        mopCS_MTopnSts = mobjFollow2(1)(1, mopCS_MTopnSts, 5)
        mopCS_MTclsSts = 0
    ElseIf mopCS_BrineSts = 2 Then
        ' LT
        mopCS_LTopnSts = mobjRampTowards(1)(1, mopCS_LTopnSts, 5, mccfUpdateTime)
        mopCS_LTclsSts = 0
        mopCS_MTopnSts = 0
        mopCS_MTclsSts = mobjRampTowards(2)(1, mopCS_MTclsSts, 3, mccfUpdateTime)
    End If
    
    If mipCS_HtrSetpt > 0 Then
        mopCS_HtrRunSts = 1
    Else
        mopCS_HtrRunSts = 0
    End If
    
    mopCS_BrineSts = mobjFollow2(2)(mipCS_BrineReq, mopCS_BrineSts, 1)
    mopP201_RunSts = mobjFollow2(3)(mipP201_RunReq, mopP201_RunSts, 2)
    mopCV_Position = mipCV_Setpt
    mopRV_Position = mipRV_Setpt
    
    If mopCS_LTopnSts = 1 Then
        mopMixBrn_Temp = mobjRampTowards(3)(mopLTBrn_Temp, mopMixBrn_Temp, 20, mccfUpdateTime)
        mopHxIn_Temp = mopHxOut_Temp - mopCV_Position / 100 * (mopHxOut_Temp - mopMixBrn_Temp) * mgvfAT_Sim_002
    ElseIf mopCS_MTopnSts = 1 Then
        mopMixBrn_Temp = mobjRampTowards(4)(mopMTBrn_Temp, mopMixBrn_Temp, 20, mccfUpdateTime)
    '    HxIn_Temp = (100 - CV_Position) / 100 * HxOut_Temp + CV_Position / 100 * MixBrn_Temp
        mopHxIn_Temp = mopHxOut_Temp - mopCV_Position / 100 * (mopHxOut_Temp - mopMixBrn_Temp) * mgvfAT_Sim_003
    End If
    
    If mipCS_HtrSetpt <> 0 Then
        ' Maximum heater power, 100 % = 300 kW
        mopHxIn_Temp = mopHxIn_Temp + (mipCS_HtrSetpt / 100) * 300 * mgvfAT_Sim_004
    End If
    
    mopHxOut_Temp = mopHxIn_Temp + (mgvfAT_Previous - mopHxIn_Temp) * mgvfAT_Sim_005
    
    ' Maximum fan power = 850 kW @ 950 rpm
    mopAirtemp = mgvfAT_Previous + (mopHxIn_Temp - mgvfAT_Previous) * 0.1 + mopMF_Speed / 950 * 850 * mgvfAT_Sim_006
    
    If mopAirtemp > 55 Then
       mopAirtemp = 55
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'    If mgvfAT_Sim = 0 Or mbResetSim Then
'
'        If mipAirtempSetpt = 0 Or mgvfAT_Sim = 0 Then
'            mgvfAirHXInletTemp = 25
'            mgvfAirtempRaw = 25
'            mgvfBrineHXInletTemp = 25
'            mgvfBrineHXOutletTemp = 25
'            mopCoolSysSupplyTemp = 11
'            mgvfBrineMixingPtTemp = 25
'            mgvfBrinePumpTemp = 25
'            mgvfBrineRecircTemp = 25
'            mgvfBrineSupplyTemp = 5
'            mgvfShellTemp1 = 25
'            mgvfShellTemp2 = 25
'
'            ' Filter time constants, ramp times
'            mgvfCoolingValveFilt = 15#  ' %/s
'            mgvfRecircValveFilt = 15#   ' %/s
'            mgvfBrineRecircTempFilt = 4#
'            mgvfBrineMixingPtTempFilt = 0#
'            mgvfAT_dT_HeaterFilt = 0#
'            mgvfBrineHXInletTempFilt = 8#
'            mgvfBrineHXOutletTempFilt = 0#
'            mgvfAirtempRawFilt = 7#
'            mgvfAT_dT_AirloadFilt = 10#
'
'            mgvfIdleWSFactor = 0.8
'        End If
'        mopBrineHeaterFault = 1
'        mopBrineSystemReady = 1
'        mopHXCoolingReady = 1
'        mopCSPLC_FCV_512_02 = 100#
'        mopCSPLC_CoolWaterTemp = 10#
'        mgvfAT_FUAMax = 343.41 ' kW/degC = FUA
'        mgvfAT_h_wall = 0.664 * 0.72 ^ 0.333 * 1000000# ^ 0.5 * 1 / 1000 ' = 0.5952 kW/m2-K
'        mgvfAT_h_insul = 0.000233 ' kW/m2-K
'        mgvfAT_areaWall = 2000 ' m2
'
'        mgvfAT_Sim = 1
'        If mbResetSim Then
'            GoTo Done
'        End If
'    End If
'
'    ' Cp = Heat capacity of air ~= 1.0035 kJ/kg/degC
'    ' Rho = Density of air is 1.225 kg/m3 (or parameter mipAirDensity)
'    ' Vol/s = Max volumetric flow rate (small nozzle) is = 250 km/h * 5 m2 * 1000 m/km / 3600 s/h = 347 m3/s
'    ' Main fan max speed = 550 rpm
'    ' Air Capacitance (kW/degC) = Vol/s (m3/s) * Rho (kg/m3) * Cp (kJ/kg-K)
'    '                           = WS * nozzlesize / 3.6 * 1.225 * 1.0035
'
''    mgvfAT_AirCap = 427 * mopMainFanSpeed / 550# * 273# / (273# + mipAirtempSetpt)
'
'    If mipNozzleSize = 1 Then
'        fNozzArea = 7#
'    Else
'        fNozzArea = 5#
'    End If
'
'    ' Idle
'
'    Select Case mipIdleStatus
'        Case 0      ' Normal - no idle
'            mgvfAT_Windspeed = mipWindspeed
'            mgvfAT_WindspeedSP = mipWindspeedSetpt
'
'        Case Else   ' 1=idle, 2=moving
'            mgvfAT_Windspeed = mgvfIdleWSFactor * mccfWindspeedSmallMax * mopMainFanSpeed / mccfMainFanSpeedMax
'            mgvfAT_WindspeedSP = mgvfAT_Windspeed
'
'    End Select
'
'    mgvfAT_AirCap = mgvfAT_Windspeed * fNozzArea / 3.6 * mipAirDensity * 1.0035 * 273# / (273# + mipAirtempSetpt)
'
'    If mgvfAT_AirCap < 50 Then
'      mgvfAT_AirCap = 50
'    End If
'
'    mgvfAT_Previous = mobjRampTowards(5)(mgvfAirtempRaw, mgvfAT_Previous, mipAirtempRate / 60, 0.01)
'
'    ' main fan max speed = 550 rpm, max power = 1600 kW
'    mgvfAT_QFan = mccfMainFanPowerMax * (mopMainFanSpeed / mccfMainFanSpeedMax) ^ 2 * (273# / (273# + mipAirtempSetpt))
'
'    mopCSPLC_ChillerPumpSts = mobjFollow(0)(mipBrineSystemControl, 5)
'
'    ' GEA are compressors for low temperatures.
'    ' HTP is a cooling tower water to brine heat exchanger for high temperatures.
'    If mopCSPLC_ChillerPumpSts = 1 Then
'        If mopCoolSysSupplyTemp >= 12 Then
'            mopCSPLC_GEAStatus = 0
'            mopCSPLC_HTPStatus = 1
'        ElseIf mopCoolSysSupplyTemp <= 10 Then
'            mopCSPLC_GEAStatus = 1
'            mopCSPLC_HTPStatus = 0
'        ElseIf mopCSPLC_GEAStatus = 0 And mopCSPLC_HTPStatus = 0 Then
'            mopCSPLC_GEAStatus = 1
'            mopCSPLC_HTPStatus = 0
'        End If
'        mopCoolSysSupplyTemp = mobjRampTowards(6)(mipBrineSupplyTempSetpt, mopCoolSysSupplyTemp, 2)
'    Else
'        mopCSPLC_GEAStatus = 0
'        mopCSPLC_HTPStatus = 0
'    End If
'
'    mopCSPLC_CoolWaterTemp = mobjRampTowards(7)(10#, mopCSPLC_CoolWaterTemp, 0.1, 0.1)  ' cooling tower water temperature
'
'    If mopCSPLC_GEAStatus = 1 Then
'        fCapacity = Ctl_Limit((mopCoolSysSupplyTemp - mipBrineSupplyTempSetpt) * 6, 5, 100)
'    Else
'        fCapacity = 0
'    End If
'
'    mopCSPLC_Comp2Capacity = mopCSPLC_Comp1Capacity
'    mopCSPLC_Comp1Capacity = mobjRampTowards(8)(fCapacity, mopCSPLC_Comp1Capacity, 0.1, 0.1)
'    If fCapacity = 0 Then mopCSPLC_Comp1Capacity = 0
'
'    mopDynoCoolingReady = mobjFollow(1)(mipDynoCoolingCmd, 5)
'    mopBrineSystemStatus = mopHXCoolingReady
'    mopBrineHeaterStatus = mobjFollow(2)(mipBrineHeaterControl, 2)
'    mopBrineHXPumpStatus = mopCSPLC_ChillerPumpSts  ' mopBrineSystemStatus
'    mopCoolSysTank3Temp = mopCoolSysSupplyTemp + 0.1
'    mopCoolSysTank2Temp = mopCoolSysSupplyTemp + 0.2
'    mopCoolSysTank1Temp = mopCoolSysSupplyTemp + 0.3
'    mgvfBrineSupplyTemp = mopCoolSysSupplyTemp + 0.1
'
'    mopCoolingValve = mobjRampTowards(9)(mipCoolingValveSetpt, mopCoolingValve, mgvfCoolingValveFilt)
'    mopRecircValve = mobjRampTowards(10)(mipRecircValveSetpt, mopRecircValve, mgvfRecircValveFilt)
'
'    mgvfBrineHXOutletTemp = Ctl_Limit(mgvfBrineHXOutletTemp, -55, 65)
'    mgvfBrineRecircTemp = mobjFilter1(0)(mgvfBrineHXOutletTemp, mgvfBrineRecircTempFilt)
'    mgvfBrineRecircTemp = Ctl_Limit(mgvfBrineRecircTemp, -55, 65)
'
'    mgvfBrineMixingPtTemp = mgvfBrineSupplyTemp * (mopCoolingValve / 100) + mgvfBrineRecircTemp * (1 - mopCoolingValve / 100)
'    mgvfBrineMixingPtTemp = mobjFilter1(1)(mgvfBrineMixingPtTemp, mgvfBrineMixingPtTempFilt)
'    mgvfBrineMixingPtTemp = Ctl_Limit(mgvfBrineMixingPtTemp, -55, 65)
'
'    'mgvfBrineMixingPtTemp = 27
'
'    If mopBrineHXPumpStatus = 1 Then
'       mgvfAT_dT_Pump = 0.2
'    Else
'       mgvfAT_dT_Pump = 0#
'    End If
'
'    If mopBrineHXPumpStatus = 1 Then
'        mgvfBrinePumpTemp = mgvfBrineMixingPtTemp + mgvfAT_dT_Pump
'    Else
'        mgvfBrinePumpTemp = mgvfBrineMixingPtTemp
'    End If
'
'    If mipBrineHeaterSetpt > 0 Then
'        ' Maximum heater power, 100 % = 1250 kW
'        ' The HX pump flows 482 m3/h of Antifrogen (brine).
'        ' The Density (g/cm3) of Antifrogen is 1.365 = 1.365 g/m3 * 100^3 cm3/m3 * (1/1000) kg/g = 1365 kg/m3
'        ' The Specific Heat Capacity (kJ/kg·K) of Antifrogen KF in water is 2.6
'        ' The heat transfer (kW) is Qdot(brine) = 482 m3/h * (1/3600)h/s * 1365 kg/m3 * 2.6 (kJ/kg-K) * dT degK
'        '                                       = 475.17 * dT
'        ' The delta-temperature is,
'        '       dT = Qdot / (Mdot * Cp)
'        '          = Qdot / (482 / 3600 * 1365 * 2.6)
'        '          = Qdot / 475.17
'
'        mgvfAT_dT_Heater = mipBrineHeaterSetpt / 100 * 1250 / 475.17
'    Else
'        mgvfAT_dT_Heater = 0
'    End If
'    mgvfAT_dT_Heater = Ctl_Limit(mgvfAT_dT_Heater, 0, 100)  ' degC
'    mgvfAT_dT_Heater = mobjFilter1(2)(mgvfAT_dT_Heater, mgvfAT_dT_HeaterFilt)
'
'    mgvfBrineHXInletTemp = mobjFilter1(3)(mgvfBrineMixingPtTemp + mgvfAT_dT_Pump + mgvfAT_dT_Heater, mgvfBrineHXInletTempFilt)
'    mgvfBrineHXInletTemp = Ctl_Limit(mgvfBrineHXInletTemp, -55, 65)
'
'    mgvfAT_MaCpa = mgvfAT_Windspeed / 3.6 * fNozzArea * mipAirDensity * 1.0035 * 273# / (273# + mipAirtempSetpt)
'
'    If mopBrineHXPumpStatus = 1 Then
'        ' mgvfAT_McCpc = 482 m3/h * (1/3600)h/s * 1365 kg/m3 * 2.6 (kJ/kg-K) * 273# / (273# + mipAirtempSetpt)
'        '              = 475.17 kJ/s-K * 273# / (273# + mipAirtempSetpt)
'        mgvfAT_McCpc = 475.17 * 273# / (273# + mipAirtempSetpt)
'    Else
'        mgvfAT_McCpc = 0#
'    End If
'
'    mgvfAirHXInletTemp = Ctl_Limit(mgvfAirHXInletTemp, -55, 65)
'
'    mgvfAT_FUA = (mgvfAT_WindspeedSP / mccfWindspeedSmallMax) ^ 0.8 * mgvfAT_FUAMax
'    ' mgvfAT_FUAMax = Heat exchanger coefficient * heat exchanger area
'    If mgvfAT_FUA <> 0 And mgvfAT_MaCpa <> 0 And mgvfAT_McCpc <> 0 Then
'        mgvfAT_Q = 2 * (mgvfAirHXInletTemp - mgvfBrineHXInletTemp) / (2 / mgvfAT_FUA + 1 / mgvfAT_MaCpa + 1 / mgvfAT_McCpc)
'    Else
'        mgvfAT_Q = 0
'    End If
'
'    If mgvfAT_McCpc <> 0# Then
'        fTDelta = mgvfAT_Q / mgvfAT_McCpc
'    Else
'        fTDelta = 0#
'    End If
'
'    mgvfBrineHXOutletTemp = mobjRampTowards(11)(mgvfBrineHXInletTemp + fTDelta, mgvfBrineHXOutletTemp, 0.1)
''    mgvfBrineHXOutletTemp = mobjFilter1(4)(mgvfBrineHXOutletTemp, mgvfBrineHXOutletTempFilt)
'    mgvfBrineHXOutletTemp = Ctl_Limit(mgvfBrineHXOutletTemp, -55, 65)
'
'    If mgvfAT_MaCpa <> 0# Then
'        ' heat transfer = f(wind speed)
'        ' mgvfAT_HXAirOut = mgvfAirHXInletTemp - (mgvfAT_WindspeedSP / mccfWindspeedSmallMax) ^ 0.8 * mgvfAT_Q / mgvfAT_MaCpa
'        fTDelta = -mgvfAT_Q / mgvfAT_MaCpa
'    Else
'        fTDelta = 0#
'    End If
'
'    mgvfAT_HXAirOut = mobjRampTowards(12)(mgvfAirHXInletTemp + fTDelta, mgvfAT_HXAirOut, 0.1)
'    mgvfAT_HXAirOut = Ctl_Limit(mgvfAT_HXAirOut, -40, 60)
'    mgvfAirtempRaw = mobjFilter1(5)(mgvfAT_HXAirOut, mgvfAirtempRawFilt)
'
'    ' The HX pump flows 482 m3/h of Antifrogen (brine).
'    ' The density of Antifrogen is 1.365 g/cm3 = 1365 kg/m3
'    ' The specific heat capacity of Antifrogen KF in water is 2.6 kJ/kg-K
'    '
'    ' The brine heat transfer is,
'    '       Qdot = Mdot * Cp * dT
'    ' where:
'    '       Qdot is power (kW)
'    '       Mdot is the mass flow rate (kg/s)
'    '       Cp is the specific heat of the fluid (kJ/kg-K)
'    '       dT is the temperature change of the fluid (K)
'    '
'    '       Qdot = 482 (m3/h) * 1/3600 (h/s) * 1365 (kg/m3) * 2.6 (kJ/kg-K) * dT (K)
'    '            = 475.17 * dT
'
'    mgvfAT_dTc_HX = mgvfBrineHXOutletTemp - mgvfBrineHXInletTemp
'    fTavg = (mgvfBrineHXOutletTemp + mgvfBrineHXInletTemp) / 2
'    mgvfAT_Qbrine = mgvfAT_McCpc * mgvfAT_dTc_HX
'
'    ' Small nozzle area is 5 m2.
'    ' Air flow = WS km/h * 1 m/s / 3.6 km/h * 5 m2
'    '          = WS (km/h) * 1.3889 (m3/s/(km/h))
'    ' Air density (Rhoair) ~= 1.225 kg/m3
'    ' Air specific heat = 1.0035 kJ/kg-K
'    ' The air heat transfer is,
'    '       Qdot = Mdot * Cp * dT
'    '       Qdot = WS (km/h) * 1.3889 (m3/s/(km/h)) * Rhoair (kg/m3) * 1.0035 (kJ/kg-K) * dT (K)
'    '            = WS * 1.3889 * Rhoair * 1.0035 * dT
'
'    mgvfAT_dTa_HX = mgvfAirHXInletTemp - mgvfAT_HXAirOut
'    fTavg = (mgvfAirHXInletTemp + mgvfAT_HXAirOut) / 2
'    mgvfAT_Qair = mgvfAT_MaCpa * mgvfAT_dTa_HX
'
'    ' LMTD:
'    '
'    ' dTA = Tai - Tco
'    ' dTB = Tao - Tci
'    '
'    ' LMTD = (dTA - dTB)/ln(dTA/dTB)        ' Logarithmic Mean Temperature Difference
'    ' Qdot = U * A * LMTD
'    ' U * A = Qdot/LMTD
'
'    fdTA = mgvfAirHXInletTemp - mgvfBrineHXOutletTemp
'    fdTB = mgvfAT_HXAirOut - mgvfBrineHXInletTemp
'
'    If fdTA = 0 Or fdTB = 0 Or (fdTA * fdTB) <= 0 Or Abs(fdTA / fdTB - 1) < 0.01 Then
'        mgvfAT_LMTD = 0
'        mgvfAT_FUA_LMTD = 0
'    Else
'        mgvfAT_LMTD = (fdTA - fdTB) / Log(fdTA / fdTB)
'        mgvfAT_FUA_LMTD = mgvfAT_Q / mgvfAT_LMTD
'    End If
'
'    ' MTD:
'    '
'    ' MTD = (Tai + Tao)/2 - (Tci + Tco)/2   ' Mean Temperature Difference
'    ' Qdot = U * A * MTD
'    ' U * A = Qdot/MTD
'
'    mgvfAT_MTD = (mgvfAirHXInletTemp + mgvfAT_HXAirOut) / 2 - (mgvfBrineHXInletTemp + mgvfBrineHXOutletTemp) / 2
'    If mgvfAT_MTD = 0 Then
'        mgvfAT_FUA_MTD = 0
'    Else
'        mgvfAT_FUA_MTD = mgvfAT_Q / mgvfAT_MTD
'    End If
'
'    mgvfAT_Qairload = Ctl_Limit(mgvfAT_Qairload, -5000, 5000)
'
'    ' Maximum fan power = 1600 kW @ 550 rpm. dT = Qdot / (Mdot * Cp)
'    If mopMainFanSpeed > 5 And mipBrineSystemControl = 1 Then     ' only if fan and cooling system are running
'        fTDelta = (mgvfAT_QFan + mipSSLampPower + mipDynoPower + mgvfAT_Qairload + mgvfAT_Qwall + mgvfAT_Qinsul) / mgvfAT_AirCap * mgvfTimeInterval_Sim
'    Else
'        fTDelta = 0
'    End If
'
'    mgvfAirHXInletTemp = mgvfAirtempRaw + fTDelta
'    mgvfAT_dT_Airload = mobjFilter1(6)((mgvfAirHXInletTemp + mgvfAT_HXAirOut) / 2 - mgvfAT_Avg, mgvfAT_dT_AirloadFilt)
'    mgvfAT_Qairload = mccfTunnelVolume * mipAirDensity * 1.006 * mgvfAT_dT_Airload
'    ' Save the average temperature for the next dt_Airload calc
'    mgvfAT_Avg = (mgvfAirHXInletTemp + mgvfAT_HXAirOut) / 2
'
'    ' Wall effect:
'    ' Aw = wall surface area ~= 2000 m2
'    ' Mw = wall mass (10mm wall) ~= 2000 m2 * 0.010 m * 7850 kg/m3 = 157000 kg
'    ' Cpw = wall heat capacity = 0.5 kJ/kg-K
'    ' hins = foam insulation heat transfer coeff ~= 2.33e-4 kW/m2-K
'    ' Dw = average wall hydraulic diameter ~= 6m
'    ' DwArea = Pi * (Dw/2)^2 = 28.27 m2
'    ' Cpa = air heat capacity ~= 1.006 kJ/kg-K
'    ' mua = air dynamic viscosity ~= 1.974e-5 kg/m-s
'    ' ka = air heat conductivity ~= 2.76e-5 kW/m-K
'    ' Rhoa = air density ~= 1.145 kg/m3 @20C
'    ' Vair = average air velocity ~= WS/3.6 m/s * DwArea / NozzleArea
'    ' Pr = Prandtl# = Cpa * mua / ka = 1.006 * 1.974e-5 / 2.76e-5 = 0.72
'    ' Re = Reynolds# = Vair (m/s) * Rhoa (kg/m3) * L (m) / mua (kg/m-s)
'    ' Vair = WS / 3.6 (m/s) * (NozzArea / HydDiamArea)
'    ' Qw-a = hw-a * Aw * (Twall - (Tao + Tai)/2)
'    ' Qins = hins * Aw * (Tamb - Twall)
'    ' Twall = (Qins - Qw-a) / (Cpw * Mw) + Twallold
'
'    mgvfAT_Re_wall = (mgvfAT_WindspeedSP / 3.6 * 5# / 28.27) * mipAirDensity * 100# / 0.00001974
'    mgvfAT_areaWall = 100# * 3.14159 * 6#
'    If mgvfAT_Re_wall > 500000# Then
'        ' Nu = Pr ^ 0.33 * (0.037 * Re ^ 0.8 -871)
'        ' h = Nu * k / L
'        mgvfAT_h_wall = 0.72 ^ 0.33 * (mgvfAT_Re_wall ^ 0.8 - 871) * 0.0000276 / 100#
'    Else
'        mgvfAT_h_wall = 0.664 * 0.72 ^ 0.333 * mgvfAT_Re_wall ^ 0.5 * 0.0000276 / 100#
'    End If
'
'    mgvfAT_Qwall = mgvfAT_h_wall * mgvfAT_areaWall * (mgvfShellTemp1 - (mgvfAirtempRaw + mgvfAirHXInletTemp) / 2)
'    mgvfAT_Qinsul = mgvfAT_h_insul * mgvfAT_areaWall * (mopOutsideTemp - mgvfShellTemp1)
'
'    mgvfShellTemp1 = (mgvfAT_Qinsul - mgvfAT_Qwall) / (0.5 * 157000) * mgvfTimeInterval_Sim + mgvfShellTemp1
'    mgvfShellTemp2 = mgvfShellTemp1
'
'    ' Truncate RTD outputs (+/-0.1C) to replicate the data the PLC provides
'    mopAirHXInletTemp = CSng(Format(mgvfAirHXInletTemp, "0.0"))
'    mopAirtempRaw = CSng(Format(mgvfAirtempRaw, "0.0"))
'    mopBrineHXInletTemp = CSng(Format(mgvfBrineHXInletTemp, "0.0"))
'    mopBrineHXOutletTemp = CSng(Format(mgvfBrineHXOutletTemp, "0.0"))
'    mopBrineMixingPtTemp = CSng(Format(mgvfBrineMixingPtTemp, "0.0"))
'    mopBrinePumpTemp = CSng(Format(mgvfBrinePumpTemp, "0.0"))
'    mopBrineRecircTemp = CSng(Format(mgvfBrineRecircTemp, "0.0"))
'    mopBrineSupplyTemp = CSng(Format(mgvfBrineSupplyTemp, "0.0"))
'    mopShellTemp1 = CSng(Format(mgvfShellTemp1, "0.0"))
'    mopShellTemp2 = CSng(Format(mgvfShellTemp2, "0.0"))
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done
    Resume
    
Done:
    Exit Sub
End Sub

Private Sub Calc_DV_Sim()

    If mgvfDV_Sim = 0 Or mbResetSim Then
        mgvfDV_Sim = 1
    
        mopCCE_Perm = 1
        mopCCE_Remote = 1
    
        mopDCA_FaultSts = 1
        mopDCA_Perm = 1
        mopDCA_Remote = 1
        mopDCA_Temp = 21
    
        mopDV_Gly_Temp = 70
    
        mopDV_Oil_Temp = 50
    
        mopDyPit_Temp = 21
    
        mopVH_IgnPerm = 1
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    mopCCE_RunSts = mobjFollow2(4)(mipCCE_Start, mopCCE_RunSts, 1)
    mopDCA_RunSts = mobjFollow2(5)(mipDCA_Start, mopDCA_RunSts, 1)
    mopVH_IgnSts = mobjFollow2(6)(mipVH_IgnEnable, mopVH_IgnSts, 1)
    
    If mopDCA_RunSts = 1 Then
        mopDCA_Temp = mobjRampTowards(13)(mipDCA_Target, mopDCA_Temp, 1, mccfUpdateTime)
    End If
    
    If mipDV_Gly_Start = 1 Then
        mopDV_Gly_Temp = mobjRampTowards(14)(mipDV_Gly_TTgt, mopDV_Gly_Temp, 1, mccfUpdateTime)
    End If
    
    If mipDV_Oil_Start = 1 Then
        mopDV_Oil_Temp = mobjRampTowards(15)(mipDV_Oil_TTgt, mopDV_Oil_Temp, 1, mccfUpdateTime)
    End If
    
    If mipDyPit_Start = 1 Then
        mopDyPit_Temp = mobjRampTowards(16)(mipDyPit_Target, mopDyPit_Temp, 1, mccfUpdateTime)
    End If
    
Done:

End Sub

Private Sub Calc_DY_Sim()

    If mgvfDY_Sim = 0 Or mbResetSim Then
        mgvfDY_Sim = 1
'        mopDy_Remote = 1
'        mopDy_CommSts = 1
        mopDy_Fault = 1
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
Done:

End Sub

Private Sub Calc_EX_Sim()
    
    Const PROC_NAME = "Calc_EX_Sim"
    
    On Error GoTo Error
    
    If mgvfEX_Sim = 0 Or mbResetSim Then
        mopEX_Fault = 1
        mopMUA_HX109Flt = 1
        mopEX_Remote = 1
        mopEX_Perm = 1
    
        mgvfEX_Sim = 1
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    mopEX_RunSts = mobjFollow2(7)(mipEX_Start, mopEX_RunSts, 3)
    
    If mopEX_RunSts = 1 Then
        mopEX_FanSpeed = mobjRampTowards(17)(mipEX_FanSetpt, mopEX_FanSpeed, 400, mccfUpdateTime)
    Else
        mopEX_FanSpeed = mobjRampTowards(18)(0, mopEX_FanSpeed, 400, mccfUpdateTime)
    End If
    
    mopBP_Raw = mopEX_FanSpeed * (-500 / 1728)
    
    'MUA_Flow = EX_Flow
    mopMUA_Temp = mopAirtemp

    
    
    
'    If mgvfEX_Sim = 0 Or mbResetSim Then
'        mopExFanFault = 1
'        mopExFanPermissive = 1
'        mopExFanInvtrRdy = 1
'        mgvfEX_Sim = 1
'        If mbResetSim Then
'            GoTo Done
'        End If
'    End If
'
'    If (mopExFanFault = 0) Or (mopExFanPermissive = 0) Then
'        mopExFanStatus = 0
'    Else
'        mopExFanStatus = mobjFollow(3)(mipExFanControl, 1)
'    End If
'
'    If mopExFanFault = 0 And mipExFanFaultReset = 1 Then
'        mopExFanFault = mobjFollow(4)(mipExFanFaultReset, 1)
'    End If
'
'    If mopExFanStatus = 0 Then
'        mopExFanSpeed = mobjRampTowards(19)(0, mopExFanSpeed, 40)
'        mopExPressureRawEast = 0#
'        mopExPressureRawWest = 0#
'    Else
'        mopExFanSpeed = mobjRampTowards(20)(mipExFanSpeedSetpt, mopExFanSpeed, 50#, 0.005)
'        If mipExPressureEnaEast = 1 Then
'            mopExPressureRawEast = mobjRampTowards(21)(-(mopExFanSpeed * 5), mopExPressureRawEast, 100#)
'        Else
'            mopExPressureRawEast = 0#
'        End If
'        If mipExPressureEnaWest = 1 Then
'            mopExPressureRawWest = mobjRampTowards(22)(-(mopExFanSpeed * 5), mopExPressureRawWest, 100#)
'        Else
'            mopExPressureRawWest = 0#
'        End If
'
'    End If
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub Calc_HD_Sim()
    
    Const PROC_NAME = "Calc_HD_Sim"
    
'    Dim fHx_AvgTemp
'    Dim fCondensTempDelta
'    Dim fPrecoolTemp
'    Dim fmccfTunnelVolume
'    Dim fTD
'    Dim fRH
'    Dim fT
'    Dim fMUATempTarget
'    Dim fMUADPTarget
'    Dim fHumidRatioTunnel
'    Dim fVolumeDelta
'    Dim fAirDensityMUA
'    Dim fKgH2OMUA
'    Dim fAirDensityTun
'    Dim fKgH2OTun
'    Dim fKgH2O
'    Dim fKgH2OMainSteam
'    Dim fKgH2OMUASteam
'    Dim fKgH2OHXCondens
'    Dim fKgH2OWallCondens
    
    On Error GoTo Error
    
    Dim fHx_AvgTemp
    Dim fCondensTemp
    
    If mgvfHD_Sim = 0 Or mbResetSim Then
    
       mopHD_Mirror = 1
       mopHD_SmplFlow = 1
    
'       mopPStatic = 101.325   'kPa
       mopSteamPress = 150#   '>45 kPa
       mopRelHumidity = 5#
    
       mopMUA_DP = 10#
       mopMUA_Temp = mopAirtemp
       mopMUA_Flow = 0#
    
       mgvfHD_Sim_001 = 0.0005
       mgvfHD_Sim_002 = 0.001
       mgvfHD_Sim_003 = -0.0015
       mgvfHD_Sim = 1
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    mopSV_Position = mobjRampTowards(23)(mipSV_Setpt, mopSV_Position, 5, mccfUpdateTime)
    
    fHx_AvgTemp = (mopHxIn_Temp + mopHxOut_Temp) / 2
    
    If fHx_AvgTemp < mipDewpoint Then
       fCondensTemp = mipDewpoint - fHx_AvgTemp
    Else
       fCondensTemp = 0#
    End If
    
    mopRelHumidity = mopRelHumidity + mopSV_Position * mgvfHD_Sim_001 - fCondensTemp * mgvfHD_Sim_002 + mgvfHD_Sim_003
    
    If mopRelHumidity > 100# Then
       mopRelHumidity = 100#
    ElseIf mopRelHumidity < 0# Then
       mopRelHumidity = 0#
    End If
    
    mopMUA_DP = mobjRampTowards(24)(mipDP_Target, mopMUA_DP, 0.5, mccfUpdateTime, 0.1)
    mopMUA_Temp = mopAirtemp
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'    If mgvfHD_Sim = 0 Or mbResetSim Then
'        If mopDewpointRaw = 0# Then
'            mopDewpointRaw = 5#  ' degC
'            mopSteamPressSupply = 515#
'            mopSteamPressMainGrid = 175#
'            mopSteamPressMUAGrid = 134#
'            mopOutsideTemp = 25#
'            mopOutsideHumidity = 35#
'        End If
'        mopMUAFault = 1
'        mopMUAReady = 1
'        mopMUARemote = 1
'
'        mgvfHD_MainSteamGridFactor = 1#       ' Main steam factor
'        mgvfHD_MUASteamGridFactor = 1#       ' MUA steam factor
'        mgvfHD_HXCondens = -0.02     ' HX condensation factor
'        mgvfHD_WallCondens = -0.02     ' Wall condensation factor
'        mgvfHD_Sim = 1
'        If mbResetSim Then
'            GoTo Done
'        End If
'    End If
'
'    ' From: http://bmcnoldy.rsmas.miami.edu/Humidity.html
'    '
'    '   RH: =100*(EXP((17.625*TD)/(243.04+TD))/EXP((17.625*T)/(243.04+T)))
'    '   TD: =243.04*(LN(RH/100)+((17.625*T)/(243.04+T)))/(17.625-LN(RH/100)-((17.625*T)/(243.04+T)))
'    '    T: =243.04*(((17.625*TD)/(243.04+TD))-LN(RH/100))/(17.625+LN(RH/100)-((17.625*TD)/(243.04+TD)))
'    '    (• replace "T", "TD", and "RH" with your actual cell references)
'    '    (• T and TD inputs/outputs to the equations are in Celsius)
'
'    fmccfTunnelVolume = 4215 ' m3
'    ' Main grid steam valve = 550 kg/h
'    ' MUA steam valve = 75 kg/h
'    mopSteamValvePos = mobjRampTowards(25)(mipSteamValveSetpt, mopSteamValvePos, 5)
'    mopMUASteamValvePos = mobjRampTowards(26)(mipMUASteamValveSetpt, mopMUASteamValvePos, 5)
'
'    fHx_AvgTemp = (mopBrineHXInletTemp + mopBrineHXOutletTemp) / 2
'
'    If fHx_AvgTemp < mopDewpointRaw Then
'       fCondensTempDelta = mopDewpointRaw - fHx_AvgTemp
'    Else
'       fCondensTempDelta = 0#
'    End If
'
'    If mopMUAFault = 0 Then
'      mopMUAStatus = 0
'    ElseIf mopMUAReady = 1 Then
'      mopMUAStatus = mobjFollow(5)(mipMUAControl, 2)
'    End If
'
'    mopMUAPreCoolStatus = mobjFollow(6)(mipMUAPreCoolControl, 3)
'    mopMUAWheelStatus = mobjFollow(7)(mipMUAWheelControl, 3)
'    fPrecoolTemp = 5# ' somewhere between +2 to +8 degC
'    mopMUAFlow = 8372 * mopMUAFanSpeed / 100# ' m3/h
'
'    If mopMUAStatus = 0 Then
'        mopMUAFanSpeed = mobjRampTowards(27)(0, mopMUAFanSpeed, 10)
'
'    Else
'        mopMUAFanSpeed = mobjRampTowards(28)(mipMUAFanSpeedSetpt, mopMUAFanSpeed, 10)
'
'        If mipMUADryingMode = 3 Then ' Full Dry
'            ' % Aftercooler temp = 5, mipDewpoint = -50
'    '        If mipMUATempSetpt < 5 Then
'    '          fMUATempTarget = mipMUATempSetpt
'    '        Else
'    '          fMUATempTarget = fPrecoolTemp
'    '        End If
'
'            fMUADPTarget = -50
'            mopMUATempAfterCool = fPrecoolTemp
'
'        ElseIf mipMUADryingMode = 2 Then ' Precooler only
'    '        fMUATempTarget = fPrecoolTemp
'            fMUADPTarget = mipOutsideDewpoint
'            mopMUATempAfterCool = fPrecoolTemp
'
'        Else ' None - unconditioned
'    '        fMUATempTarget = mopOutsideTemp
'            fMUADPTarget = mipOutsideDewpoint
'            mopMUATempAfterCool = mopOutsideTemp
'
'        End If
'
'        fMUATempTarget = mipMUATempSetpt
'    End If
'
'    fT = mobjRampTowards(29)(CSng(fMUATempTarget), mopMUATemperature, 0.2)
'    mgvfHD_MUADewpoint = mobjRampTowards(30)(CSng(fMUADPTarget), mgvfHD_MUADewpoint, 0.2)
'    fTD = mgvfHD_MUADewpoint
'    fRH = 100 * (Exp((17.625 * fTD) / (243.04 + fTD)) / Exp((17.625 * fT) / (243.04 + fT)))
'    mopMUARelHumidity = fRH
'    mopMUATemperature = fT
'    mopMUADewpoint = Ctl_DewPointTemp(mopMUATemperature, mopMUARelHumidity)
'
'    mgvfHD_HumidRatioMUA = Ctl_AbsHumidity(mopMUADewpoint, mipPabs)
'    fHumidRatioTunnel = Ctl_AbsHumidity(mipDewpoint, mipPabs)
'
'    fVolumeDelta = (mopMUAFlow / 3600#) * mgvfTimeInterval_Sim
'    fAirDensityMUA = Ctl_DryAirDensity(mopMUADewpoint, mopMUATemperature, mipPabs)
'    fKgH2OMUA = fVolumeDelta * fAirDensityMUA * mgvfHD_HumidRatioMUA ' m3 * kgair/m3 * kgH2O/kgair = kgH2O
'    fAirDensityTun = Ctl_DryAirDensity(mipDewpoint, mipAirtemp, mipPabs)
'    fKgH2OTun = (fmccfTunnelVolume - fVolumeDelta) * fAirDensityTun * fHumidRatioTunnel
'    fKgH2OMainSteam = mgvfHD_MainSteamGridFactor * mopSteamValvePos * 550 / 3600 * mgvfTimeInterval_Sim
'    fKgH2OMUASteam = mgvfHD_MUASteamGridFactor * mopMUASteamValvePos * 75 / 3600 * mgvfTimeInterval_Sim
'    fKgH2OHXCondens = mgvfHD_HXCondens * fCondensTempDelta
'    fKgH2OWallCondens = mgvfHD_WallCondens
'
'    fKgH2O = fKgH2OTun + fKgH2OMUA + fKgH2OMainSteam + fKgH2OMUASteam + fKgH2OHXCondens + fKgH2OWallCondens
'
'    mgvfHD_HumidRatioTunnel = mobjFilter1(7)(fKgH2O / (fmccfTunnelVolume * fAirDensityTun), 5)
'    mgvfHD_HumidRatioTunnel = Ctl_Limit(mgvfHD_HumidRatioTunnel, 0.000015, 0.1)
'    mopHumidityRatio = mgvfHD_HumidRatioTunnel * 1000
'
'    If mipPabs = 0 Then
'        mipPabs = 101#
'    End If
'    mopDewpointRaw = Ctl_AbsHumidityToDewPoint(mipPabs, mgvfHD_HumidRatioTunnel, mipAirtemp)  ' Dewpoint_HumidityRatio1(mipPabs, mgvfHD_HumidRatioTunnel, mipAirtemp)
'    mopDewpointRaw = Ctl_Limit(mopDewpointRaw, -55, mipAirtemp)
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub Calc_HR_Sim()
    
    Const PROC_NAME = "Calc_HR_Sim"
'    Const HR_TEMP_RATE = 0.5    ' degC/s
'    Const HR_TEMP_NOISE = 0.1   ' degC
    Dim fHR_RunSts As Single
    Dim fHR_HxInTmp As Single
    Dim fHR_HxOutTmp As Single
    Dim fHR_CVSetpt As Single
    Dim fHR_HtrSetpt As Single
    Dim fHR_PnlTemp As Single
    Dim i As Long
    
    On Error GoTo Error
    
    If mgvfHR_Sim = 0 Or mbResetSim Then
        mgvfHR_Sim = 1
    
        mopHR_FaultSts = 1
        mopHR_Perm = 1
        mopHR_EstopSts = 1
    
        mopHR1_FaultSts = 1
        mopHR2_FaultSts = 1
        mopHR3_FaultSts = 1
        mopHR4_FaultSts = 1
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    If mipHR_RunSts = 1 Then
        mopHR_Brn_Temp = mobjRampTowards(31)(mipHR_BrnSetpt, mopHR_Brn_Temp, 0.5, mccfUpdateTime)
    End If
    
    For i = 1 To 4
        Select Case i
            Case 1
                fHR_RunSts = mipHR1_RunSts
                fHR_HxInTmp = mopHR1_HxInTmp
                fHR_HxOutTmp = mopHR1_HxOutTmp
                fHR_CVSetpt = mipHR1_CVSetpt
                fHR_HtrSetpt = mipHR1_HtrSetpt
                fHR_PnlTemp = mopHR1_PnlTemp
            Case 2
                fHR_RunSts = mipHR2_RunSts
                fHR_HxInTmp = mopHR2_HxInTmp
                fHR_HxOutTmp = mopHR2_HxOutTmp
                fHR_CVSetpt = mipHR2_CVSetpt
                fHR_HtrSetpt = mipHR2_HtrSetpt
                fHR_PnlTemp = mopHR2_PnlTemp
            Case 3
                fHR_RunSts = mipHR3_RunSts
                fHR_HxInTmp = mopHR3_HxInTmp
                fHR_HxOutTmp = mopHR3_HxOutTmp
                fHR_CVSetpt = mipHR3_CVSetpt
                fHR_HtrSetpt = mipHR3_HtrSetpt
                fHR_PnlTemp = mopHR3_PnlTemp
            Case 4
                fHR_RunSts = mipHR4_RunSts
                fHR_HxInTmp = mopHR4_HxInTmp
                fHR_HxOutTmp = mopHR4_HxOutTmp
                fHR_CVSetpt = mipHR4_CVSetpt
                fHR_HtrSetpt = mipHR4_HtrSetpt
                fHR_PnlTemp = mopHR4_PnlTemp
        End Select
        
        If fHR_RunSts = 1 Then
            fHR_HxInTmp = fHR_HxOutTmp - fHR_CVSetpt / 100 * (fHR_HxOutTmp - mopHR_Brn_Temp)
            If fHR_HtrSetpt <> 0 Then
                ' Maximum heater power, 100 % = 21 kW
                fHR_HxInTmp = fHR_HxInTmp + (fHR_HtrSetpt / 100) * 21
             End If
        
            fHR_HxOutTmp = fHR_HxInTmp + (fHR_PnlTemp - fHR_HxInTmp)
            fHR_PnlTemp = fHR_PnlTemp + (fHR_HxInTmp - fHR_PnlTemp) * 0.1
        
            If fHR_PnlTemp > 80 Then
                fHR_PnlTemp = 80
            ElseIf fHR_PnlTemp < 10 Then
                fHR_PnlTemp = 10
            End If
        Else
            fHR_HxInTmp = mobjRampTowards(32)(mopAirtemp, fHR_HxInTmp, 0.5, mccfUpdateTime)
            fHR_HxOutTmp = fHR_HxInTmp
            fHR_PnlTemp = fHR_HxInTmp
        End If
    
        Select Case i
            Case 1
                mopHR1_HxInTmp = fHR_HxInTmp
                mopHR1_HxOutTmp = fHR_HxOutTmp
                mopHR1_PnlTemp = fHR_PnlTemp
            Case 2
                mopHR2_HxInTmp = fHR_HxInTmp
                mopHR2_HxOutTmp = fHR_HxOutTmp
                mopHR2_PnlTemp = fHR_PnlTemp
            Case 3
                mopHR3_HxInTmp = fHR_HxInTmp
                mopHR3_HxOutTmp = fHR_HxOutTmp
                mopHR3_PnlTemp = fHR_PnlTemp
            Case 4
                mopHR4_HxInTmp = fHR_HxInTmp
                mopHR4_HxOutTmp = fHR_HxOutTmp
                mopHR4_PnlTemp = fHR_PnlTemp
        End Select
    Next i
    
'    If HR1_RunSts = 1 Then
'        HR1_HxInTmp = HR1_HxOutTmp - HR1_CVSetpt / 100 * (HR1_HxOutTmp - HR_Brn_Temp)
'        If HR1_HtrSetpt <> 0 Then
'            ' Maximum heater power, 100 % = 21 kW
'            HR1_HxInTmp = HR1_HxInTmp + (HR1_HtrSetpt / 100) * 21
'         End If
'
'        HR1_HxOutTmp = HR1_HxInTmp + (HR1_PnlTemp - HR1_HxInTmp)
'        HR1_PnlTemp = HR1_PnlTemp + (HR1_HxInTmp - HR1_PnlTemp) * 0.1
'
'        If HR1_PnlTemp > 80 Then
'            HR1_PnlTemp = 80
'        ElseIf HR1_PnlTemp < 10 Then
'            HR1_PnlTemp = 10
'        End If
'    Else
'        HR1_HxInTmp = RampTowards(Airtemp, HR1_HxInTmp, 0.5)
'        HR1_HxOutTmp = HR1_HxInTmp
'        HR1_PnlTemp = HR1_HxInTmp
'    End If
'
'    If HR2_RunSts = 1 Then
'        HR2_HxInTmp = HR2_HxOutTmp - HR2_CVSetpt / 100 * (HR2_HxOutTmp - HR_Brn_Temp)
'        If HR2_HtrSetpt <> 0 Then
'            ' Maximum heater power, 100 % = 21 kW
'            HR2_HxInTmp = HR2_HxInTmp + (HR2_HtrSetpt / 100) * 21
'         End If
'
'        HR2_HxOutTmp = HR2_HxInTmp + (HR2_PnlTemp - HR2_HxInTmp)
'        HR2_PnlTemp = HR2_PnlTemp + (HR2_HxInTmp - HR2_PnlTemp) * 0.1
'
'        If HR2_PnlTemp > 80 Then
'            HR2_PnlTemp = 80
'        ElseIf HR2_PnlTemp < 10 Then
'            HR2_PnlTemp = 10
'        End If
'    Else
'        HR2_HxInTmp = RampTowards(Airtemp, HR2_HxInTmp, 0.5)
'        HR2_HxOutTmp = HR2_HxInTmp
'        HR2_PnlTemp = HR2_HxInTmp
'    End If
'
'    If HR3_RunSts = 1 Then
'        HR3_HxInTmp = HR3_HxOutTmp - HR3_CVSetpt / 100 * (HR3_HxOutTmp - HR_Brn_Temp)
'        If HR3_HtrSetpt <> 0 Then
'            ' Maximum heater power, 100 % = 21 kW
'            HR3_HxInTmp = HR3_HxInTmp + (HR3_HtrSetpt / 100) * 21
'         End If
'
'        HR3_HxOutTmp = HR3_HxInTmp + (HR3_PnlTemp - HR3_HxInTmp)
'        HR3_PnlTemp = HR3_PnlTemp + (HR3_HxInTmp - HR3_PnlTemp) * 0.1
'
'        If HR3_PnlTemp > 80 Then
'            HR3_PnlTemp = 80
'        ElseIf HR3_PnlTemp < 10 Then
'            HR3_PnlTemp = 10
'        End If
'    Else
'        HR3_HxInTmp = RampTowards(Airtemp, HR3_HxInTmp, 0.5)
'        HR3_HxOutTmp = HR3_HxInTmp
'        HR3_PnlTemp = HR3_HxInTmp
'    End If
'
'    If HR4_RunSts = 1 Then
'        HR4_HxInTmp = HR4_HxOutTmp - HR4_CVSetpt / 100 * (HR4_HxOutTmp - HR_Brn_Temp)
'        If HR4_HtrSetpt <> 0 Then
'            ' Maximum heater power, 100 % = 21 kW
'            HR4_HxInTmp = HR4_HxInTmp + (HR4_HtrSetpt / 100) * 21
'         End If
'
'        HR4_HxOutTmp = HR4_HxInTmp + (HR4_PnlTemp - HR4_HxInTmp)
'        HR4_PnlTemp = HR4_PnlTemp + (HR4_HxInTmp - HR4_PnlTemp) * 0.1
'
'        If HR4_PnlTemp > 80 Then
'            HR4_PnlTemp = 80
'        ElseIf HR4_PnlTemp < 10 Then
'            HR4_PnlTemp = 10
'        End If
'    Else
'        HR4_HxInTmp = RampTowards(Airtemp, HR4_HxInTmp, 0.5)
'        HR4_HxOutTmp = HR4_HxInTmp
'        HR4_PnlTemp = HR4_HxInTmp
'    End If

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'    If mgvfHR_Sim = 0 Or mbResetSim Then
'        mopHRController01Fault = 138
'        mopHRController02Fault = 138
'        mopHRController03Fault = 138
'
'        mopHRController01Status = 62
'        mopHRController02Status = 62
'        mopHRController03Status = 62
'
'        mopHRPanel01AnaErr = 61
'        mopHRPanel02AnaErr = 61
'        mopHRPanel03AnaErr = 61
'        mopHRPanel04AnaErr = 61
'        mopHRPanel05AnaErr = 61
'
'        mopHRPanel01AnaErrTC = 61
'        mopHRPanel02AnaErrTC = 61
'        mopHRPanel03AnaErrTC = 61
'        mopHRPanel04AnaErrTC = 61
'        mopHRPanel05AnaErrTC = 61
'
'        mopHRPanel01LimitState = 61
'        mopHRPanel02LimitState = 61
'        mopHRPanel03LimitState = 61
'        mopHRPanel04LimitState = 61
'        mopHRPanel05LimitState = 61
'
'        mopHRPanel01LimitStatus = 1667
'        mopHRPanel02LimitStatus = 1667
'        mopHRPanel03LimitStatus = 1667
'        mopHRPanel04LimitStatus = 1667
'        mopHRPanel05LimitStatus = 1667
'
'        mgvfHR_Sim = 1
'        If mbResetSim Then
'            GoTo Done
'        End If
'    End If
'
'    If mipHRController01Start = 63 Then
'        mopHRController01Status = 63
'    End If
'    If mipHRController01Stop = 63 Then
'        mopHRController01Status = 62
'    End If
'
'    If mipHRController02Start = 63 Then
'        mopHRController02Status = 63
'    End If
'    If mipHRController02Stop = 63 Then
'        mopHRController02Status = 62
'    End If
'
'    If mipHRController03Start = 63 Then
'        mopHRController03Status = 63
'    End If
'    If mipHRController03Stop = 63 Then
'        mopHRController03Status = 62
'    End If
'
'    If mipHRPanel01Control = 1 Then
'        mopHRPanel01Temp = mobjRampTowards(33)(mipHRPanel01Setpt, mopHRPanel01Temp, HR_TEMP_RATE, HR_TEMP_NOISE)
'    Else
'        mopHRPanel01Temp = mobjRampTowards(34)(mipAirtemp, mopHRPanel01Temp, HR_TEMP_RATE, HR_TEMP_NOISE)
'    End If
'    mopHRPanel01TempTC = mopHRPanel01Temp
'
'    If mipHRPanel02Control = 1 Then
'        mopHRPanel02Temp = mobjRampTowards(35)(mipHRPanel02Setpt, mopHRPanel02Temp, HR_TEMP_RATE, HR_TEMP_NOISE)
'    Else
'        mopHRPanel02Temp = mobjRampTowards(36)(mipAirtemp, mopHRPanel02Temp, HR_TEMP_RATE, HR_TEMP_NOISE)
'    End If
'    mopHRPanel02TempTC = mopHRPanel02Temp
'
'    If mipHRPanel03Control = 1 Then
'        mopHRPanel03Temp = mobjRampTowards(37)(mipHRPanel03Setpt, mopHRPanel03Temp, HR_TEMP_RATE, HR_TEMP_NOISE)
'    Else
'        mopHRPanel03Temp = mobjRampTowards(38)(mipAirtemp, mopHRPanel03Temp, HR_TEMP_RATE, HR_TEMP_NOISE)
'    End If
'    mopHRPanel03TempTC = mopHRPanel03Temp
'
'    If mipHRPanel04Control = 1 Then
'        mopHRPanel04Temp = mobjRampTowards(39)(mipHRPanel04Setpt, mopHRPanel04Temp, HR_TEMP_RATE, HR_TEMP_NOISE)
'    Else
'        mopHRPanel04Temp = mobjRampTowards(40)(mipAirtemp, mopHRPanel04Temp, HR_TEMP_RATE, HR_TEMP_NOISE)
'    End If
'    mopHRPanel04TempTC = mopHRPanel04Temp
'
'    If mipHRPanel05Control = 1 Then
'        mopHRPanel05Temp = mobjRampTowards(41)(mipHRPanel05Setpt, mopHRPanel05Temp, HR_TEMP_RATE, HR_TEMP_NOISE)
'    Else
'        mopHRPanel05Temp = mobjRampTowards(42)(mipAirtemp, mopHRPanel05Temp, HR_TEMP_RATE, HR_TEMP_NOISE)
'    End If
'    mopHRPanel05TempTC = mopHRPanel05Temp
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub Calc_IC_Sim()
    
    Const PROC_NAME = "Calc_IC_Sim"
    
    On Error GoTo Error
    
    If mgvfIdle_Sim = 0 Or mbResetSim Then
    
        mgvfIdle_Sim = 1
    
        mopIDL_Fault = 1
        mopIDL_Perm = 1
        mopIDL_Remote = 1
    
        mopIDL_ModeOn = 0
        mopIDL_ModeOff = 1
    
        mopIDL_FlOpdSts = 0
        mopIDL_FlCldSts = 1
    
        mopIDL_ShCldSts = 0
        mopIDL_ShOpdSts = 1
    
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    If mipIDL_Start = 1 Then
        mopIDL_ModeOn = 1
        mopIDL_ModeOff = 0
        If mopIDL_FlOpdSts = 0 Then
            mopIDL_FlOpdSts = mobjFollow2(8)(1, mopIDL_FlOpdSts, 5)
            mopIDL_FlCldSts = 0
        ElseIf mopIDL_ShCldSts = 0 Then
            mopIDL_ShCldSts = mobjFollow2(9)(1, mopIDL_ShCldSts, 5)
            mopIDL_ShOpdSts = 0
        End If
    Else
        If mopIDL_ShOpdSts = 0 Then
            mopIDL_ShOpdSts = mobjFollow2(10)(1, mopIDL_ShOpdSts, 5)
            mopIDL_ShCldSts = 0
        ElseIf mopIDL_FlCldSts = 0 Then
            mopIDL_FlCldSts = mobjFollow2(11)(1, mopIDL_FlCldSts, 5)
            mopIDL_FlOpdSts = 0
        Else
            mopIDL_ModeOn = 0
            mopIDL_ModeOff = 1
        End If
    End If

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'    Dim bIdleClosedStatus
'    Dim bIdleOpenStatus
'    Dim bIdleInterStatus
'
'    If mgvfIC_Sim = 0 Or mbResetSim Then
'        mopIdlePermissive = 1
'        mopIdleInMotion = 0
'        mopIdleDriveFault = 1
'
'        mgvfIC_Sim = 1
'        If mbResetSim Then
'            GoTo Done
'        End If
'    End If
'
'    bIdleClosedStatus = mobjFollow(8)(mipIdleCloseControl, 5#)
'    bIdleOpenStatus = mobjFollow(9)(mipIdleOpenControl, 5#)
'    bIdleInterStatus = mobjFollow(10)(mipIdleInterControl, 5#)
'
'    If mipIdleCloseControl = 1 Then
'       mopIdleOpenStatus = 0: mopIdleInterStatus = 0
'    ElseIf mipIdleOpenControl = 1 Then
'       mopIdleClosedStatus = 0: mopIdleInterStatus = 0
'    ElseIf mipIdleInterControl = 1 Then
'       mopIdleClosedStatus = 0: mopIdleOpenStatus = 0
'    End If
'
'    If bIdleClosedStatus = 1 Then
'       mopIdleClosedStatus = 1: mopIdleOpenStatus = 0: mopIdleInterStatus = 0
'    ElseIf bIdleOpenStatus = 1 Then
'       mopIdleClosedStatus = 0: mopIdleOpenStatus = 1: mopIdleInterStatus = 0
'    ElseIf bIdleInterStatus = 1 Then
'       mopIdleClosedStatus = 0: mopIdleOpenStatus = 0: mopIdleInterStatus = 1
'    End If
'
'    If mopIdlePermissive = 1 Then
'       If mipIdleCloseControl = 1 Then
'          mopIdleOpenStatus = 0
'          mopMainFanSpeedLim1 = 0
'          If mopIdleClosedStatus = 0 Then
'             mopIdleInMotion = 1
'          Else
'             mopIdleInMotion = 0
'          End If
'       End If
'       If mipIdleOpenControl = 1 Or mipIdleInterControl = 1 Then
'          mopIdleClosedStatus = 0
'          If mopIdleOpenStatus = 0 And mopIdleInterStatus = 0 Then
'             mopIdleInMotion = 1
'          Else
'             mopIdleInMotion = 0
'             mopMainFanSpeedLim1 = 1
'          End If
'       End If
'    Else
'       mopIdleInMotion = 0
'    End If

    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub Calc_Faults_Sim()
    
    Const PROC_NAME = "Calc_Faults_Sim"
    
    On Error GoTo Error
    
    If mgvfFaultsSim = 0 Or mbResetSim Then
    
'        mopBLFanFault = 1
'        mopBLFanPerm = 1
'
'        mopEpbControlConsole = 1
'        mopEpbTS_ControlRoom = 1
'        mopEpbTS_PlantRoom = 1
'        mopEpbTS_Diffuser = 1
'        mopEpbVehicle = 1
'        mopEStopStatus = 1
'        mopEpbSummSts = 1
'        mopEpbHXEast = 1
'        mopEpbHXWest = 1
'        mopEpbMainFan = 1
'        mopEpbCorner4East = 1
'        mopEpbCorner4West = 1
'        mopEpbCorner3 = 1
'        mopEpbCorner2 = 1
'        mopEpbCorner2West = 1
'        mopEpbVehicleEntry = 1
'
'        mopExFanFault = 1
'        mopExFanPermissive = 1
'
'        mopDrCorner2 = 1
'        mopDrCorner3 = 1
'        mopDrCorner4East = 1
'        mopDrCorner4West = 1
'        mopDrDynoBasement = 1
'        mopDrHXInsideEast1 = 1
'        mopDrHXInsideEast2 = 1
'        mopDrHXInsideWest = 1
'        mopDrHXOutsideEast = 1
'        mopDrHXOutsideWest = 1
'        mopDrMainFan = 1
'        mopDrSummSts = 1
'        mopDrTS_ControlRoom = 1
'        mopDrTS_ControlRoomMagLockSts = 1
'        mopDrTS_Diffuser = 1
'        mopDrTS_IC_East = 1
'        mopDrTS_IC_West = 1
'        mopDrTS_Plant = 1
'        mopDrVehAccessEast = 1
'        mopDrVehAccessEastMagLockSts = 1
'        mopDrVehAccessWest = 1
'        mopDrVehAccessWestMagLockSts = 1
'
'        'mopHXCoolingReady = 1
'
'        mopLockoutKey = 1
'        mopKeySwSummSts = 1
'
'        mopFireSts = 1
'        mopFireSummSts = 1
'
'        mopHG_CO2TSFlt = 1
'        mopHG_COTSFlt = 1
'        mopHG_H2TSFlt = 1
'        mopHG_HCTSFlt = 1
'        mopHG_NOxTSFlt = 1
'        mopHG_CO2DyFlt = 1
'        mopHG_CODyFlt = 1
'        mopHG_HCDyFlt = 1
'        mopHG_NOxDyFlt = 1
'
'        mopIdleDriveFault = 1
'
'        mopMainFanVibFault = 1
'        mopMainFanWindingTempFault = 1
'
'        mopMUAFault = 1
'        mopMUAReady = 1
'
'        mopOpSwitch = 8     ' Testing
'
'        mopPurgeFanFault = 1
'
'        mopPrsHxHi = 1
'        mopPrsPlenumHi = 1
'        mopPrsPlenumHiHi = 1
'        mopPrsPlenumLo = 1
'        mopPrsPlenumLoLo = 1
'
'        mopSnowGunDriveBrakingOK = 1
'        mopSnowGunDriveFault = 1
'
'        mopTunnelFault = 1
'
'        mopCSPLC_Status = 1
'        mopHRPLC_Status = 1
'        mopMFPLC_Status = 1
'        mopPRPLC_Status = 1
'        mopSFPLC_Status = 1
'        mopSNPLC_Status = 1
        
        mgvfFaultsSim = 1
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
'    If mopPrsHxHi = 1 And mopPrsPlenumHiHi = 1 And mopPrsPlenumLoLo = 1 Then
'        mopPressSummSts = 1
'    Else
'        mopPressSummSts = 0
'    End If
'
'    mopHG_TSSummSts = mopHG_CO2TSFlt And mopHG_COTSFlt And mopHG_H2TSFlt _
'                        And mopHG_HCTSFlt And mopHG_NOxTSFlt
'    mopHG_DySummSts = mopHG_CO2DyFlt And mopHG_CODyFlt And mopHG_HCDyFlt _
'                        And mopHG_NOxDyFlt
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
    
End Sub

Private Sub Calc_MF_Sim()
    
    Const PROC_NAME = "Calc_MF_Sim"
    Dim fMfRate
    
    On Error GoTo Error
    
    'mgvfMF_Previous = (0.03) * mipMF_Setpt + (1 - 0.03) * mopMF_Speed
    
    
    
    
    
    
    
    
    
    
    
    
    
'    If mgvfMF_Sim = 0 Or mbResetSim Then
'        mopMainFanFault1 = 1
'        mopMainFanFault2 = 1
'        mopMainFanDriveFault1 = 1
'        mopMainFanDriveFault2 = 1
'        mopMainFanRemote = 1
'        mopMainFanPermissive = 1
'        mopEStopStatus = 1
'        mipNozzleSize = 1
'
'        mopMainFanSpeedLim1 = 1
'        mopMainFanSpeedLim2 = 1
'
'        mopMainFanSpeedLim1Status = 1
'        mopMainFanSpeedLim2Status = 1
'
'        mopMainFanWindingTempFault = 1
'        mopMainFanVibFault = 1
'
'        mgvfMF_Sim = 1
'        If mbResetSim Then
'            GoTo Done
'        End If
'    End If
'
'    mopMFCoolingReady = mobjFollow(11)(mipMainFanCoolingCmd, 5)
'
'    If (mopMainFanFault1 = 0) Or (mopMainFanFault2 = 0) Or (mopMainFanDriveFault1 = 0) Or (mopMainFanDriveFault2 = 0) Then
'      mopMainFanStatus = 0
'    Else
'      mopMainFanStatus = mobjFollow(12)(mipMainFanControl, 5)
'    End If
'
'    If mipMainFanFastStopCmd = 1 Then
'      fMfRate = 100
'    Else
'      fMfRate = 55
'    End If
'
'    If mopMainFanStatus = 0 Then
'      mopMainFanSpeed = mobjRampTowards(43)(0, mopMainFanSpeed, CSng(fMfRate))
'    Else
'      mopMainFanSpeed = mobjRampTowards(44)(mipMainFanSpeedSetpt, mopMainFanSpeed, CSng(fMfRate))
'    End If
'
'    mopMainFanPower = (mopMainFanSpeed / mccfMainFanSpeedMax) ^ 2 * mccfMainFanPowerMax
'
'    mopMainFanSpeedLim1Status = mobjFollow(13)(mopMainFanSpeedLim1, 5)
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub Calc_PG_Sim()

    Const PROC_NAME = "Calc_PG_Sim"
    
    On Error GoTo Error
    
    If mgvfPG_Sim = 0 Or mbResetSim Then
        mopPG_FanFault = 1
        mopPG_FanPerm = 1
        mopPG_FanRemote = 1
    
        mopPG_HLActive = 1
    
        mopPGDy_FanFlt = 1
        mopPGDy_FanPerm = 1
        mopPGDy_Remote = 1
    
        mgvfPG_Sim = 1
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    If mipPG_HL_Start = 1 Then
        If mipPG_HL_DelRst = 1 Then
            mopPG_HL_Tmr = 30
        Else
            mopPG_HL_Tmr = mobjRampTowards(45)(0, mopPG_HL_Tmr, 1, mccfUpdateTime)
        End If
        
        If mopPG_HL_Tmr > 0 Then
            ' Delay start until timer elapses
        Else
            mopPG_FanRunH = mobjFollow2(12)(mipPG_HL_Start, mopPG_FanRunH, 1)
        End If
    Else
        mopPG_HL_Tmr = 30
        mopPG_FanRunH = mobjFollow2(13)(mipPG_HL_Start, mopPG_FanRunH, 1)
    End If
    
    mopPG_HLRunSts = mopPG_FanRunH
    
    mopPG_FanRunL = mobjFollow2(14)(mipPG_LL_Start, mopPG_FanRunL, 1)
    
    If mipPGDy_HL_Strt = 1 Then
        If mipPGDy_HL_DRst = 1 Then
            mopPGDy_HL_Tmr = 30
        Else
            mopPGDy_HL_Tmr = mobjRampTowards(46)(0, mopPGDy_HL_Tmr, 1, mccfUpdateTime)
        End If
        
        If mopPGDy_HL_Tmr > 0 Then
            ' Delay start until timer elapses
        Else
            mopPGDy_FRunH = mobjFollow2(15)(mipPGDy_HL_Strt, mopPGDy_FRunH, 1)
        End If
    Else
        mopPGDy_HL_Tmr = 30
        mopPGDy_FRunH = mobjFollow2(16)(mipPGDy_HL_Strt, mopPGDy_FRunH, 1)
    End If
    
    mopPGDy_HLRnSts = mopPGDy_FRunH
    
    mopPGDy_FRunL = mobjFollow2(17)(mipPGDy_LL_Strt, mopPGDy_FRunL, 1)

    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub Calc_SF_Sim()

    Const PROC_NAME = "Calc_SF_Sim"
    
    On Error GoTo Error
    
    If mgvfSF_Sim = 0 Or mbResetSim Then
        mgvfSF_Sim = 1
    
        mopCtrl_Power = 1
     '   Door_vPrCl = 1
     '   Door_UpLg_BF = 1
        mopDoor_UpLg_AS = 1
     '   Door_sMTV_Cl = 1
     '   Door_OuThCl = 1
     '   Door_nMTV_Cl = 1
     '   Door_InThCl = 1
     '   Door_HX_Vest = 1
     '   Door_HX_UpSt = 1
     '   Door_HX_DnSt = 1
        mopDoor_Dy = 1
     '   Door_Diffusr = 1
        mopKey_Veh_Door = 1
     '   Key_Passthru = 1
        mopKey_Dy_Door = 1
        mopFire_Sum = 1
        mopFire_DYSup = 1
        mopLoLo_Plen_P = 1
        mopLo_Plen_P = 1
    '    Lo_DP_Pl_Dy = 1
    '    Lo_DP_Dy_Pl = 1
    '    Hi_DP_HX = 1
        mopHX_HiDeltaP = 1
    '    EPB_Setl_Ch = 1
        mopEPB_TS_NW = 1
        mopEPB_TS_SW = 1
        mopEPB_UpLeg_AS = 1
    '    EPB_UpLeg_BF = 1
        mopEStop_SumSts = 1
        mopEPB_Ctrl_Rm = 1
    '    TS_HC_Alarm = 1
    '    TS_HC_Warn = 1
    '    TS_CO_Alarm = 1
    '    TS_CO_Warn = 1
    '    R507_Alarm = 1
    '    R507_Warn = 1
    '    HC_Alarm_Sum = 1
    '    HC_Warn_Sum = 1
        mopFaultBus = 1
    '    Dy_HC_Alarm = 1
    '    Dy_HC_Warn = 1
        mopDy_Estop_Sts = 1
    '    DY_CO_Alarm = 1
    '    DY_CO_Warn = 1
    '    CO_Alarm_Sum = 1
    '    CO_Warn_Sum = 1
        mopGas_Sys_Sts = 1
    '    GasSysFloSts = 1
        
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub Calc_SR_Sim()

    Const PROC_NAME = "Calc_SR_Sim"
    
    On Error GoTo Error
    
    If mgvfSR_Sim = 0 Or mbResetSim Then
        mgvfSR_Sim = 1
        mopAL1_FaultSts = 1
        mopAL2_FaultSts = 1
        mopSR1_FaultSts = 1
        mopSR2_FaultSts = 1
        mopSR3_FaultSts = 1
        mopSR4_FaultSts = 1
        mopSR5_FaultSts = 1
        mopSR6_FaultSts = 1
        mopTA1_FaultSts = 1
        mopTA2_FaultSts = 1
    
        mopSR1_Master = 2
        mopSR2_Master = mopSR1_Master
        mopSR3_Master = mopSR1_Master
        mopSR4_Master = mopSR1_Master
        mopSR5_Master = mopSR1_Master
        mopSR6_Master = mopSR1_Master
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    mopSR1_SetptFbk = mipSR1_Setpt
    mopSR2_SetptFbk = mipSR2_Setpt
    mopSR3_SetptFbk = mipSR3_Setpt
    mopSR4_SetptFbk = mipSR4_Setpt
    mopSR5_SetptFbk = mipSR5_Setpt
    mopSR6_SetptFbk = mipSR6_Setpt
    
    mopSR1_AirTemp = mobjRampTowards(47)(mipSR1_Setpt, mopSR1_AirTemp, 2, mccfUpdateTime)
    mopSR2_AirTemp = mobjRampTowards(48)(mipSR2_Setpt, mopSR2_AirTemp, 2, mccfUpdateTime)
    mopSR3_AirTemp = mobjRampTowards(49)(mipSR3_Setpt, mopSR3_AirTemp, 2, mccfUpdateTime)
    mopSR4_AirTemp = mobjRampTowards(50)(mipSR4_Setpt, mopSR4_AirTemp, 2, mccfUpdateTime)
    mopSR5_AirTemp = mobjRampTowards(51)(mipSR5_Setpt, mopSR5_AirTemp, 2, mccfUpdateTime)
    mopSR6_AirTemp = mobjRampTowards(52)(mipSR6_Setpt, mopSR6_AirTemp, 2, mccfUpdateTime)
    
    mopSR1_RunSts = mobjFollow2(18)(mipSR1_RunReq, mopSR1_RunSts, 1)
    mopSR2_RunSts = mobjFollow2(19)(mipSR2_RunReq, mopSR2_RunSts, 1)
    mopSR3_RunSts = mobjFollow2(20)(mipSR3_RunReq, mopSR3_RunSts, 1)
    mopSR4_RunSts = mobjFollow2(21)(mipSR4_RunReq, mopSR4_RunSts, 1)
    mopSR5_RunSts = mobjFollow2(22)(mipSR5_RunReq, mopSR5_RunSts, 1)
    mopSR6_RunSts = mobjFollow2(23)(mipSR6_RunReq, mopSR6_RunSts, 1)
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub Calc_SS_Sim()
    
    Const PROC_NAME = "Calc_SS_Sim"
'    Dim i As Long
'    Dim j As Long
'    Dim fSS_LinePowerIdxCmd(1 To 7)
'    Dim fSSBasePower(1 To 3)
    
    On Error GoTo Error
    
    If mgvfSS_Sim = 0 Or mbResetSim Then
        mgvfSS_Sim = 1
        mopSS_ArmRetSts = 1
        mopSS_Fan1Fault = 1
        mopSS_Fan1Rem = 1
        mopSS_Fan2Fault = 1
        mopSS_Fan2Rem = 1
        mopSS_FanPerm = 1
        mopSS_HiTemp = 1
        mopSS_Perm = 1
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    mopSS_ArmRetSts = mobjFollow2(24)(-mipSS_ArmExtReq + 1, mopSS_ArmRetSts, 5)
    mopSS_ArmExtSts = mobjFollow2(25)(mipSS_ArmExtReq, mopSS_ArmExtSts, 5)
    
    If mipSS_ArmExtReq = 0 Then
        mopSS_ArmExtSts = 0
    Else
        mopSS_ArmRetSts = 0
    End If
    
    ' Add offset (-28.4) to InputPyro to simulate a difference - re: Compensate/Uncompensate
    mopSS_InputPyro = mipSS_InvCoef0 - 28.4 + mipSS_InvCoef1 * mipSS_Power + mipSS_InvCoef2 * mipSS_Power ^ 2 + mipSS_InvCoef3 * mipSS_Power ^ 3 + mipSS_InvCoef4 * mipSS_Power ^ 4
    
    mopSS_Fn1RunSts = mobjRampTowards(53)(mipSS_FanStart, mopSS_Fn1RunSts, 1, mccfUpdateTime)
    mopSS_Fn2RunSts = mobjRampTowards(54)(mipSS_FanStart, mopSS_Fn2RunSts, 1, mccfUpdateTime)
    mopSS_RunSts = mobjRampTowards(55)(mipSS_RunReq, mopSS_RunSts, 1, mccfUpdateTime)
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'    If mgvfSS_Sim = 0 Or mbResetSim Then
'        mopSSReady = 1
'        mopSSFault = 1
'
'        mgvfSS_Sim = 1
'        If mbResetSim Then
'            GoTo Done
'        End If
'    End If
'
'    mopSSRemoteStatus = mobjFollow(14)(mipSSRemoteRequest, 2)
'
'    mopSSHMM12Setpoint = mobjFollow(15)(mipSSHMM12SetpointCmd, 1)
'    mopSSHMM12Actual = mobjRampTowards(56)(mopSSHMM12Setpoint, mopSSHMM12Actual, 2)
'
'    mopSSHMM14Setpoint = mobjFollow(16)(mipSSHMM14SetpointCmd, 1)
'    mopSSHMM14Actual = mobjRampTowards(57)(mopSSHMM14Setpoint, mopSSHMM14Actual, 2)
'
'    mopSSHMM15Setpoint = mobjFollow(17)(mipSSHMM15SetpointCmd, 1)
'    mopSSHMM15Actual = mobjRampTowards(58)(mopSSHMM15Setpoint, mopSSHMM15Actual, 2)
'
'    mopSSHMM17Setpoint = mobjFollow(18)(mipSSHMM17SetpointCmd, 1)
'    mopSSHMM17Actual = mobjRampTowards(59)(mopSSHMM17Setpoint, mopSSHMM17Actual, 2)
'
'    mopSSWMM21mSetpoint = mobjFollow(19)(mipSSWMM21SetpointCmd, 1)
'    mopSSWMM21mActual = mobjRampTowards(60)(mopSSWMM21mSetpoint, mopSSWMM21mActual, 5)
'
'    mopSSWMM22mSetpoint = mobjFollow(20)(mipSSWMM22SetpointCmd, 1)
'    mopSSWMM22mActual = mobjRampTowards(61)(mopSSWMM22mSetpoint, mopSSWMM22mActual, 5)
'
'    mopSSWMM24mSetpoint = mobjFollow(21)(mipSSWMM24SetpointCmd, 1)
'    mopSSWMM24mActual = mobjRampTowards(62)(mopSSWMM24mSetpoint, mopSSWMM24mActual, 5)
'
'    mopSSWMM25mSetpoint = mobjFollow(22)(mipSSWMM25SetpointCmd, 1)
'    mopSSWMM25mActual = mobjRampTowards(63)(mopSSWMM25mSetpoint, mopSSWMM25mActual, 5)
'
'    mopSSWMM27mSetpoint = mobjFollow(23)(mipSSWMM27SetpointCmd, 1)
'    mopSSWMM27mActual = mobjRampTowards(64)(mopSSWMM27mSetpoint, mopSSWMM27mActual, 5)
'
'    mopSSWMM28mSetpoint = mobjFollow(24)(mipSSWMM28SetpointCmd, 1)
'    mopSSWMM28mActual = mobjRampTowards(65)(mopSSWMM28mSetpoint, mopSSWMM28mActual, 5)
'
'    mopSSIntensity1SetptFB = mobjFollow(25)(mipSSIntensity1Setpt, 1)
'    mopSSIntensity1 = mobjRampTowards(66)(mopSSIntensity1SetptFB, mopSSIntensity1, 50)
'
'    mopSSIntensity2SetptFB = mobjFollow(26)(mipSSIntensity2Setpt, 1)
'    mopSSIntensity2 = mobjRampTowards(67)(mopSSIntensity2SetptFB, mopSSIntensity2, 50)
'
'    mopSSIntensity3SetptFB = mobjFollow(27)(mipSSIntensity3Setpt, 1)
'    mopSSIntensity3 = mobjRampTowards(68)(mopSSIntensity3SetptFB, mopSSIntensity3, 50)
'
'    If mipSS_SysStartStop = 2 Then
'        mopSSStatus = 1
'    ElseIf mipSS_SysStartStop = 1 Then
'        mopSSStatus = 0
'    End If
'
'    Select Case mipSSMode
'        Case smManual, smAutoProfile
'            mopSSIntensity1 = 0
'            mopSSIntensity2 = 0
'            mopSSIntensity3 = 0
'
'            mopSSBasePower1 = mobjFollow(28)(mipSSBasePower1Setpt, 1)
'            mopSSBasePower2 = mobjFollow(29)(mipSSBasePower2Setpt, 1)
'            mopSSBasePower3 = mobjFollow(30)(mipSSBasePower3Setpt, 1)
'        Case smAutoPLC
'            If mipSSProfileIntensity > 0 Then
'                mopSSBasePower1 = mipSSIntensity1Setpt / 1200# * 100#
'                mopSSBasePower2 = mipSSIntensity2Setpt / 1200# * 100#
'                mopSSBasePower3 = mipSSIntensity3Setpt / 1200# * 100#
'            End If
'    End Select
'
'    fSS_LinePowerIdxCmd(1) = mipSS_Line1PowerIdxCmd
'    fSS_LinePowerIdxCmd(2) = mipSS_Line2PowerIdxCmd
'    fSS_LinePowerIdxCmd(3) = mipSS_Line3PowerIdxCmd
'    fSS_LinePowerIdxCmd(4) = mipSS_Line4PowerIdxCmd
'    fSS_LinePowerIdxCmd(5) = mipSS_Line5PowerIdxCmd
'    fSS_LinePowerIdxCmd(6) = mipSS_Line6PowerIdxCmd
'    fSS_LinePowerIdxCmd(7) = mipSS_Line7PowerIdxCmd
'
'    fSSBasePower(1) = mopSSBasePower1
'    fSSBasePower(2) = mopSSBasePower2
'    fSSBasePower(3) = mopSSBasePower3
'
'    For i = 1 To 28
'        With mudtSSLamp(i)
'            If .Control = 1 Then
'                .Status = 8
'            Else
'                .Status = 0
'            End If
'
'            .Offset = .OffsetSp
'            .SelectedStatus = .Selected
'
'            j = Int((i - 1) / 4) + 1
'            If fSS_LinePowerIdxCmd(j) > 0 Then
'                .Power = Ctl_Limit(fSSBasePower(fSS_LinePowerIdxCmd(j)) + .Offset, 0#, 100#)
'            Else
'                .Power = 0
'            End If
'
'            If .CloudShutterCmd = 4 Then
'                .CloudShutterStatus = 0
'            ElseIf .CloudShutterCmd = 2 Then
'                .CloudShutterStatus = 1
'            End If
'
'            If .TunnelShutterCmd = 4 Then
'                .TunnelShutterStatus = 0
'            ElseIf .TunnelShutterCmd = 2 Then
'                .TunnelShutterStatus = 1
'            End If
'
'        End With
'    Next i
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub Calc_TimeInterval_Sim()
    
    Const PROC_NAME = "Calc_TimeInterval_Sim"
    
    Dim CurrentSeconds
    
    On Error GoTo Error
    
    CurrentSeconds = Timer
    mgvfTimeInterval_Sim = CurrentSeconds - mgvfTimeInterval_Sim_001
    mgvfTimeInterval_Sim = Ctl_Limit(mgvfTimeInterval_Sim, 0.05, 1#)
    mgvfTimeInterval_Sim_001 = CurrentSeconds
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub Calc_WS_Sim()
    
    Const PROC_NAME = "Calc_WS_Sim"
    
    Dim i As Long
    Dim fWindspeed As Single
    
    On Error GoTo Error
    
    If mgvfWS_Sim = 0 Or mbResetSim Then
        mopMF_Remote = 1
        mopMF_Fault = 1
        mopMF_Perm = 1
        mopMF_SupplySts = 1
    
        mopMF_Warning = 1
    
        mopMF_SpdLim1 = 0
        mopMF_SpdLim2 = 0
        
        mgvfMF_Sim_001 = 0#     ' main fan speed filter
        mgvfMF_Sim_002 = 0.2    ' main fan speed delay filter
        mgvfQ_Sim_001 = 1#      ' dynamic pressure correction
        mgvfDP_Sim_001 = 1#     ' compressible flow factor
        mgvfDP_Sim_002 = 0#     ' filter time
        mgvfWS_Sim_001 = 0.2
        
        mopPabs = 101.325       ' kPa
        mgvfWS_Sim = 1
        
        If mbResetSim Then
            GoTo Done
        End If
    End If
    
    Select Case mgvfMF_SpeedDelayType
        Case 1
            ' MF speed "pipeline" delay
            ' - for simulating a multi-second response delay
            '   by the Amelec analog limiter
        
            mfMFSpeedSPBuffer(mlMFSpeedSPBufferIndex) = mipMF_Setpt
            
            mlMFSpeedSPBufferIndex = mlMFSpeedSPBufferIndex + 1
        
            If mlMFSpeedSPBufferIndex > mlMFSpeedSPBufferSize Then
                mlMFSpeedSPBufferIndex = 0
            End If
        
            mopMF_Speed = mfMFSpeedSPBuffer(mlMFSpeedSPBufferIndex)
        
        Case Else
            mopMF_Speed = mobjFilter1(8)(mipMF_Setpt, mgvfMF_Sim_001)
    
    End Select
    
    mgvfPstatic = mopPabs * 1000    '  Pa
       
    ' Windspeed (m/s) = MF speed (rpm) / nomFSmax (rpm) * nomWSmax (kph) / 3.6 (m/s)/kph
    
    mgvfMF_SpeedDelayed = mobjFilter1(9)(mopMF_Speed, mgvfMF_Sim_002)
    fWindspeed = mgvfMF_SpeedDelayed / 1000 * 230 / 3.6     ' rpm / rpm * kph / (m/s)/kph
'    fWindspeed = mobjFilter1(10)(mipWS_Setpt / 3.6, mgvfWS_Sim_001)
    
    ' Calculate the measured deltaP based upon the Windspeed to simulate
    '
    '   Q = 1/2 * rho * v^2
    
    mgvfQ_Sim = 0.5 * mipAirDensity * fWindspeed ^ 2    ' Q in Pa
    mgvfQ_Sim = mgvfQ_Sim * mgvfQ_Sim_001

    ' Compressible flow correction:
    '   Blockage = 1.0
    '   Pabs = 101.325 kPa
    '   Kp ~ 0, Kq ~ 1
    '
    '   Pt/Ps = (1 + (g-1)/2 * M^2)^g/(g-1)     (1.1) standard compressible flow equation
    '   M^2 = v^2 / (gRT)                       (1.2) Mach number definition
    '   Q = 1/2 * rho * v^2                     (1.3) dynamic pressure definition
    '   Ps = rho * R * T                        (1.4) ideal gas equation
    '   g = 1.4                                 (1.5) gamma
    '   g/(g-1) = 1.4/(1.4 - 1) = 3.5           (1.6) often used ...
    '
    ' So,
    '   v^2 = 2Q / rho                          (1.7) from (1.3)
    '   RT = Ps / rho                           (1.8) from (1.4)
    '   M^2 = (2Q / rho) / (g * (Ps / rho))     (1.9) substituting into (1.2)
    '       = 2Q / (g * Ps)
    '
    '   ' In (1.1), substitute for M^2, multiply both sides by Ps, subtract Ps from both sides
    '
    '   Pt - Ps = Ps * (1 + (g-1)/2 * (2Q / (g * Ps)))^g/(g-1) - Ps
    '           = Ps * [(1 + (g-1)/2 * (2Q / (g * Ps)))^g/(g-1) - 1]
    '           = Ps * [(1 + Q / (3.5 * Ps))) ^ 3.5 - 1]
    
    mgvfDP_Sim = mgvfPstatic * (((1# + mgvfQ_Sim / (3.5 * mgvfPstatic))) ^ 3.5 - 1#)
    mgvfDP_Sim = mgvfDP_Sim * mgvfDP_Sim_001
    mgvfDP_Sim = mobjFilter1(11)(mgvfDP_Sim, mgvfDP_Sim_002)
               
    mopDeltaP_Lo = Ctl_Limit(mgvfDP_Sim, 0#, mccfDPxOverHi)
    mopDeltaP_Hi = Ctl_Limit(mgvfDP_Sim, 0#, 2500#)
    
    If mipMF_Start = 0 Then
       mopMF_RunSts = mobjFollow2(26)(mipMF_Start, mopMF_RunSts, 0.5)
    Else
       mopMF_RunSts = mobjFollow2(27)(mipMF_Start, mopMF_RunSts, 5)
    End If
    
    GoTo Done
    
Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If

Done:
    Exit Sub
End Sub

Private Sub CalcDynamicPressurePa(DeltaPPa As Single, AbsPkPa As Single, Kp As Single, Kq As Single, BlockageFactor As Single, _
    CalculatedPStatickPa As Single, CalculatedQPa As Single)
'
    Const PROC_NAME = "CalcDynamicPressurePa"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Const K1 As Double = 0.285714285
    Dim PStatic As Double           ' Pa
    Dim Pt As Double                ' Pa
    
    On Error GoTo Error

    PStatic = (AbsPkPa * 1000#) + Kp * DeltaPPa
    Pt = PStatic + Kq * DeltaPPa
        
    CalculatedQPa = 1 / K1 * PStatic * ((Pt / PStatic) ^ K1 - 1#) * BlockageFactor
    If CalculatedQPa < 0 Then
        CalculatedQPa = 0#
    End If
    CalculatedPStatickPa = PStatic / 1000#
    
    GoTo Done
Error:
    CalculatedQPa = 0#
    CalculatedPStatickPa = AbsPkPa
    Resume Done
Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError lRaiseErrorNumber, TypeName(Me), PROC_NAME, sRaiseErrorText
    End If
    Exit Sub
End Sub

Public Sub CheckInterlock(Opcode, Optional CheckValue)

'Checks the specified interlock. If the interlock does not exist, then
' no error is raised.

    Const PROC_NAME = "CheckInterlock"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim sIndex As String
    Dim objInterlock As Interlock

    On Error GoTo Error
    bRaiseError = False
    
    On Error Resume Next
    sIndex = Opcode
    Set objInterlock = mobjInterlocks(sIndex)
    If Err Then
        GoTo Done
    End If
    
    mvInterlockCheckValue = CheckValue
    
    On Error GoTo Error
    objInterlock.Check

    GoTo Done
Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    lRaiseErrorNumber = Err.Number
    Resume Done
Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, TypeName(Me) & "." & PROC_NAME, sRaiseErrorText
    End If
    Exit Sub
End Sub

Private Sub Class_Initialize()

    Const PROC_NAME = "Class_Initialize"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim sClassName As String
    
    On Error GoTo Error

    bRaiseError = False
    
    On Error Resume Next

' Stop this instance if already running
    If App.PrevInstance Then
        If App.StartMode <> vbSModeStandalone Then
            EH.AppError = StdTalentStrings.stsEHInvalidActiveXLaunch
            EH.Item(1) = PROCESS_NAME
            EH.Module = TypeName(Me)
            EH.Procedure = PROC_NAME
            Call EH_MsgBox
        End If
        End
    End If
    
    On Error GoTo Error
        
    App.StartLogging "", vbLogAuto
    App.OleServerBusyRaiseError = True
    App.OleServerBusyTimeout = 1000
    
    Process.Title = PROCESS_NAME
    Process.EHCaption = PROCESS_NAME
'    If Not IsWithinIDE() Then
        Process.EHMode = ehLogOnly
'    Else
'        Process.EHMode = ehDisplayAndLog
'    End If
    Process.EHMsgsToLog = ehCriticalExclamation
    Process.LogMode = logToTalentAndFile
    
    On Error GoTo Error_CreatingObjects
    
    sClassName = "LoopTimer"
    Set tmrMain = New LoopTimer
    If App.StartMode <> vbSModeStandalone Then
        Set tmrRunOnce = New ccrpTimers6.ccrpTimer
    End If
        
    On Error GoTo Error
        
    If App.StartMode <> vbSModeStandalone Then
        tmrRunOnce.Interval = 500
        tmrRunOnce.EventType = TimerOneShot
        tmrRunOnce.Enabled = True
    End If
        
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Error_CreatingObjects:
    sRaiseErrorText = LoadResString(ridCantCreateObject, sClassName) & Err.Description
'    sRaiseErrorText = "Can't create object '" & sClassName & "': " & Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    EH.Item(1) = sClassName
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Sub

End Sub

Private Sub Class_Terminate()

    Dim i As Long
    
    On Error Resume Next
    
    tmrMain.Enabled = False
    Set tmrMain.Notify = Nothing
    Set tmrMain = Nothing
'    mobjInterlocks.Quit
'    Set mobjInterlocks = Nothing
    For i = 1 To mlNumFilter1
        Set mobjFilter1(i) = Nothing
    Next i
    For i = 1 To mlNumRampTowards
        Set mobjRampTowards(i) = Nothing
    Next i
    For i = 1 To mlNumFollow
        Set mobjFollow(i) = Nothing
    Next i
    For i = 1 To mlNumFollow2
        Set mobjFollow2(i) = Nothing
    Next i
    Set tmrElapsed = Nothing
    Set mobjCommands = Nothing
    Set mobjConstants = Nothing
    Set mobjGlobalVars = Nothing
    Set mobjEvents = Nothing
    Set mobjStats = Nothing
    Set mobjStatuses = Nothing
    Set mcolInvalidOpcodes = Nothing
    Set mobjSPGConstants = Nothing
    Set mobjSPGDataItems = Nothing
    mobjParManager.Release
    Set mobjParManager = Nothing
    mbInitialized = False
    Exit Sub

End Sub

Public Property Get Commands() As CommandListLib.Commands

' Returns the command list

' Returns
' --------------------------------------------------------
' The return value is an object ...

    Const PROC_NAME = "Commands (Get)"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    Set Commands = mobjCommands

    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Property

End Property

Public Property Get ControlConstants() As ControlLib.ControlConstants

' Returns the collection of control constants for use by an
' external debugger (Look)

' Returns
' --------------------------------------------------------
' The return value is a ControlConstants object

    Const PROC_NAME = "ControlConstants (Get)"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False
    Set ControlConstants = mobjConstants

    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Property

End Property

Friend Sub ControlLogic()

' This routine is executed at every control interval
    
    Const PROC_NAME = "ControlLogic"
    Static bInitialized As Boolean
    Dim i As Long
    Dim sCommand As String
    Dim vReceiveBuffer As Variant

    On Error GoTo Error

    ExecuteCommands
    
    If mbInitialized Then
        mobjParManager.ReadInputs
        
        If mgvfAT_Sim = 0 Or mgvfResetSim = 0 Then
            mbResetSim = True
        End If
        
        Calc_TimeInterval_Sim
        
        Calc_AT_Sim
        Calc_DV_Sim
        Calc_DY_Sim
        Calc_EX_Sim
        Calc_Faults_Sim
        Calc_HD_Sim
        Calc_HR_Sim
        Calc_IC_Sim
        Calc_MF_Sim
        Calc_PG_Sim
        Calc_SF_Sim
        Calc_SR_Sim
        Calc_SS_Sim
        Calc_WS_Sim
        
        If mbResetSim Then
            mbResetSim = False
            mgvfResetSim = 1
        End If
        
        mobjParManager.WriteOutputs
    End If
    
    If Not mobjStatuses Is Nothing Then
        mobjStatuses.FlushMessages
    End If
    
    GoTo Done

Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Suspend
    Resume Next
    Resume

Done:
    On Error Resume Next

    Exit Sub

End Sub

Public Property Get Events() As Events
    Set Events = mobjEvents
End Property

Private Sub ExecuteCommands()

' Executes the next command on the list
    
    Const PROC_NAME = "ExecuteCommands"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Static ReEntered As Boolean
    Dim objCommand As CommandListLib.Command
    Dim i As Long
    Dim objEH As EH
    Dim bFirstTime As Boolean
    
    On Error GoTo Error

    bRaiseError = False
    If ReEntered Then
        Exit Sub
    End If
    ReEntered = True

    If mobjCommands.Count = 0 Then
        GoTo Done
    End If

    Set objCommand = mobjCommands(1)
    With objCommand
        Select Case .Opcode

' Initialize
        Case ocInitialize
            Me.Initialize , objCommand
            .Remove
            mobjStatuses(STS_SIM_INITIALIZED).SetValue mbInitialized
            
        Case ocSuspend
            .Messages.Flush
            mbInitialized = False
            .Remove
            mobjStatuses(STS_SIM_INITIALIZED).SetValue mbInitialized
            SetProcessPriority priNormal
'            If IsWithinIDE() Then
'                Me.Shutdown
'            End If
        
        Case ocShutdown
            .Messages.Flush
            .Remove
            Me.Shutdown
            
        Case Else       ' All other commands require initialized state
            
            CheckInterlock ocInitialize
            
            Select Case .Opcode
            Case Else
                .Abort "Unrecognized opcode: " & .Opcode
                
            End Select
        End Select
        
    End With
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    lRaiseErrorNumber = CCommandError(Err.Number)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume ErrorExit
    Resume
    
ErrorCreatingSocket:
    sRaiseErrorText = Err.Description
    lRaiseErrorNumber = CCommandError(Err.Number)
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    Resume ErrorExit
    
ErrorExit:
    On Error Resume Next
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, objCommand.Text, sRaiseErrorText)
    objCommand.Abort sRaiseErrorText, cmdAborted
    Resume Done

Done:
    On Error Resume Next
    ReEntered = False
    Exit Sub

End Sub

Public Function GetInvalidOpcodes(Opcode) As Variant

    Dim sIndex As String
    On Error Resume Next
    sIndex = Opcode
    GetInvalidOpcodes = mcolInvalidOpcodes(sIndex)
    If Err Then
        GetInvalidOpcodes = 0
    End If
    Exit Function

End Function

Public Property Get GlobalVars() As ControlLib.GlobalVars

' Returns the collection of global variables for use by an
' external debugger (Look)

' Returns
' --------------------------------------------------------
' The return value is a GlobalVars object

    Const PROC_NAME = "GlobalVars (Get)"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False
    Set GlobalVars = mobjGlobalVars
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Property

End Property

Private Sub ICcrpTimerNotify_Timer(ByVal Milliseconds As Long)
    
    Dim lElapsed As Long
    
    On Error Resume Next
    
    ControlLogic
    
    If mbKeepStats And mbInitialized Then
        lElapsed = tmrElapsed.Elapsed
        tmrElapsed.Reset
        mobjStats.AddValue CDbl(lElapsed / 1000)
    End If
    DoEvents
    
    Exit Sub
    
End Sub

Friend Sub Initialize(Optional InitType As InitTypes = initAll, _
            Optional Command As Command)

' Perform basic initialization/startup functions

    Const PROC_NAME = "Initialize"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim bStatus As Boolean, sText As String
    Dim sFileName As String
    Dim objEH As EH
    Dim bExists As Boolean
    Dim sPath As String
    Dim sDefaultsFileName As String
    Dim i As Long
    Dim objConstants As ControlLib.ControlConstants
    
    On Error GoTo Error

    bRaiseError = False

    Select Case InitType
    Case initConstants, initAll
        bStatus = True
        sPath = GetDefaultCCFolder(App.Path)
        sFileName = sPath & PROCESS_NAME & ".csv"
        sDefaultsFileName = App.Path & "\" & App.EXEName & "Defaults.csv"
        With mobjConstants
            .OpenFileWithDefaults sFileName, sDefaultsFileName
            
            .AddString "Priority", mccsPriority, bStatus, sText
            .AddSingle "UpdateInterval", mccfUpdateTime, bStatus, sText

' >>> Beginning of [Control Constants] .AddSingle - do not move or change this line
            .AddSingle "FanDelay", mccfFanDelay, bStatus, sText
            .AddSingle "TunnelVolume", mccfTunnelVolume, bStatus, sText
' <<< End of [Control Constants] .AddSingle - do not move or change this line
            
            If bStatus = False Then
                bRaiseError = True
                sRaiseErrorText = "Can't find the following control constants in " & sFileName & ": " & sText
                lRaiseErrorNumber = tceCommandFailed
                If Command Is Nothing Then
                Else
                    Command.SuppressLogging = True
                End If
                Set objEH = New EH
                With objEH
                    .AppError = LoadResString(StdTalentStrings.stsEHCantLoadConstants)
                    .Item(1) = sText
                    .FileName = sFileName
                    .Module = TypeName(Me)
                    .Procedure = PROC_NAME
                    LogMessage2 objEH
                End With
            End If
        
            With mobjSPGConstants
                .CreatePropertyByPosition(ccUpdateTime, bExists).Value = mccfUpdateTime
            End With
        
        End With
        
        ' extra constants from other files
    
        Set objConstants = New ControlLib.ControlConstants
        With objConstants
            .OpenFile sPath & "WindSpeed.csv"
            .AddSingle "KP1_C0", mccfKp1Coeff(0), bStatus, sText
            .AddSingle "KP1_C1", mccfKp1Coeff(1), bStatus, sText
            .AddSingle "KP1_C2", mccfKp1Coeff(2), bStatus, sText
            .AddSingle "KP2_C0", mccfKp2Coeff(0), bStatus, sText
            .AddSingle "KP2_C1", mccfKp2Coeff(1), bStatus, sText
            .AddSingle "KP2_C2", mccfKp2Coeff(2), bStatus, sText
            .AddSingle "KP_DPTRANS", mccfKpDpTrans, bStatus, sText
            .AddSingle "KQ1_C0", mccfKq1Coeff(0), bStatus, sText
            .AddSingle "KQ1_C1", mccfKq1Coeff(1), bStatus, sText
            .AddSingle "KQ1_C2", mccfKq1Coeff(2), bStatus, sText
            .AddSingle "KQ2_C0", mccfKq2Coeff(0), bStatus, sText
            .AddSingle "KQ2_C1", mccfKq2Coeff(1), bStatus, sText
            .AddSingle "KQ2_C2", mccfKq2Coeff(2), bStatus, sText
            .AddSingle "KQ_DPTRANS", mccfKqDpTrans, bStatus, sText
            .AddSingle "DPXHI", mccfDPxOverHi, bStatus, sText
            .AddSingle "DPXLO", mccfDPxOverLo, bStatus, sText
        End With
        Set objConstants = Nothing
               
    End Select
    
    
    mlDefPriority = GetControllerPriority(mccsPriority)
    mbResetSim = True
    
    If mgvfListParsMissing = 0 Then
        On Error Resume Next
        mobjParManager.RefreshObjects
        On Error GoTo Error
        mbInitialized = True
    Else
        Select Case InitType
        Case initAll, initPars
            On Error Resume Next
            mobjParManager.RefreshObjects
            If Err Then
                lRaiseErrorNumber = CCommandError(Err.Number)
                sRaiseErrorText = Err.Description
                Set objEH = New EH
                With objEH
                    Set .VBErr = Err
                    .AppError = LoadResString(StdTalentStrings.stsEHInitializeFailedParsMissing)
                    .Item(1) = Command.Text
                    .Module = TypeName(Me)
                    .Procedure = PROC_NAME
                    If Command Is Nothing Then
                    Else
                        .Item(1) = Command.Text
                    End If
                    LogMessage2 objEH
                    If Command Is Nothing Then
                    Else
                        Command.SuppressLogging = True
                    End If
                End With
                GoTo Done
            End If
            mbInitialized = mobjParManager.Ready
        End Select
    End If
    
    For i = 0 To mlNumFilter1
        Set mobjFilter1(i) = New Filter1
    Next i
    
    For i = 0 To mlNumRampTowards
        Set mobjRampTowards(i) = New Ramp1
    Next i
    
    For i = 0 To mlNumFollow
        Set mobjFollow(i) = New Follow
    Next i
    
    For i = 0 To mlNumFollow2
        Set mobjFollow2(i) = New Follow2
    Next i
    
    ' Fan Delay pipe line.
    mlMFSpeedSPBufferSize = CLng(mccfFanDelay / mccfUpdateTime)
    ReDim mfMFSpeedSPBuffer(mlMFSpeedSPBufferSize)
    
    For i = 0 To mlMFSpeedSPBufferSize
        mfMFSpeedSPBuffer(i) = mipMF_Setpt
        mlMFSpeedSPBufferIndex = 0
    Next i
    
    GoTo Done

Error:
    mbInitialized = False
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        If Command Is Nothing Then
        Else
            EH.Item(1) = Command.Text
        End If
        Call EH_MsgBox
        If Command Is Nothing Then
        Else
            Command.SuppressLogging = True
        End If
    End If
    Resume Done
    Resume

Done:
    On Error Resume Next
    mobjStatuses(STS_SIM_INITIALIZED).SetValue mbInitialized
    mobjConstants.CloseFile
    Set objEH = Nothing
    If mbInitialized Then
        SetProcessPriority mlDefPriority
    Else
        SetProcessPriority priNormal
    End If
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(lRaiseErrorNumber, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Sub

End Sub

Public Property Get Initialized() As Boolean
Attribute Initialized.VB_Description = "Returns whether or not the controller is initialized (controlling) or not"

' Returns whether or not the controller is initialized (controlling) or
' not

' Returns
' --------------------------------------------------------
' True
' False

    Const PROC_NAME = "Initialized (Get)"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    Initialized = mbInitialized

    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError OLEERR_PROPERTY_GET_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText
    End If

    Exit Property

End Property

Public Property Get Interlocks() As Interlocks
    Set Interlocks = mobjInterlocks
End Property

Private Sub LogMessageEH(MessageText As String, SendingRoutine As String, Optional Items, Optional MessageType As MessageTypes = msgInformation)

    Dim objEH As EH
    Dim sItems() As String
    Dim i As Long
    
    On Error Resume Next
    Set objEH = New EH
    With objEH
        .AppError = MessageText
        If Not IsMissing(Items) Then
            For i = 1 To 10
                Err.Clear
                .Item(i) = Items(i)
                If Err Then
                    Exit For
                End If
            Next i
        End If
        .Module = TypeName(Me)
        .Procedure = SendingRoutine
        .MessageType = MessageType
        LogMessage2 objEH
    End With
    Exit Sub
    
End Sub

Sub Main()

' This is the main startup routine
    
    Const PROC_NAME = "Main"
    Dim bRunning As Boolean
    
    On Error Resume Next
' Stop this instance if already running
    If App.PrevInstance Then
        If App.StartMode <> vbSModeStandalone Then
            EH.AppError = AE.InvalidActiveXLaunch
            EH.Item(1) = PROCESS_NAME
            EH.Module = TypeName(Me)
            EH.Procedure = PROC_NAME
            Call EH_MsgBox
        End If
        End
    End If
    
    On Error GoTo Error
    App.StartLogging "", vbLogAuto
    App.OleServerBusyRaiseError = True
    App.OleServerBusyTimeout = 1000
    
    Process.Title = PROCESS_NAME
    Process.EHCaption = PROCESS_NAME
    'If Not IsWithinIDE() Then
        Process.EHMode = ehLogOnly
    'Else
    '    Process.EHMode = ehDisplayAndLog
    'End If
    Process.EHMsgsToLog = ehCriticalExclamation
    Process.LogMode = logToTalentAndFile
                
    GoTo Done
    
Error:
    If Err <> 364 Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume ErrorExit

ErrorExit:
    EH.AppError = ERR_CANT_START
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    On Error Resume Next
    GoTo Done

Done:
    Exit Sub
    
End Sub

Public Property Get ParManager() As Object
    Set ParManager = mobjParManager
End Property

Public Property Get Ready() As Boolean

'
'
'


' Returns
' --------------------------------------------------------
' True
' False
 
    On Error Resume Next
    Ready = mbReady
    Exit Property

End Property

Public Sub Shutdown()

'
'
'


    Const PROC_NAME = "Shutdown"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim objSOM As SOM
    
    On Error GoTo Error

    bRaiseError = False
                
    SetProcessPriority priNormal
    
' Shut down actions
' ------------------
    
' Clear any pending commands
    mbReady = False
    mobjCommands.Clear "Controller shutting down"
    DoEvents
        
    mbInitialized = False
    mobjParManager.ReleaseOutputs
    
    mobjEvents.RaiseQuit
        
    DoEvents
    Unload mobjForm
    Set mobjForm = Nothing
    
    Set objSOM = New SOM
    objSOM.RemoveObject SOM_CONTROLLER
    Set objSOM = Nothing
    
    Class_Terminate
    
    Pause 2#
    
    If IsWithinIDE() Then
        End
    Else
        TerminateProcess GetCurrentProcess(), 0
    End If
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    Resume Next

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Sub

End Sub

Private Sub Start()
Attribute Start.VB_Description = "Private start routine"

' Private start routine

    Const PROC_NAME = "Start"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim i As Long, objProcess As Object
    Dim sClassName As String
    Dim objProcessManagerProcess As Object
    Dim objSPGM As SharedPropertyGroupManager
    Dim bExists As Boolean
    
    On Error GoTo Error

    bRaiseError = False
        
    On Error GoTo Error_CreatingObjects
        
    sClassName = "CommandListLib.Commands"
    Set mobjCommands = New CommandListLib.Commands

    sClassName = "ControlLib.ControlConstants"
    Set mobjConstants = New ControlLib.ControlConstants

    sClassName = "Events"
    Set mobjEvents = New Events

    sClassName = "Interlocks"
    Set mobjInterlocks = New Interlocks

    sClassName = "Invalid Opcodes Collection"
    Set mcolInvalidOpcodes = New Collection
            
    sClassName = "frmControl"
    Set mobjForm = New frmControl
    
    sClassName = "SharedPropertyGroupManager"
    Set objSPGM = New SharedPropertyGroupManager
    
'    sClassName = "SharedPropertyGroup - Constants"
    Set mobjSPGConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists)
'    Set mobjSPGDataItems = objSPGM.CreatePropertyGroup(SPG_DATA, 0, 0, bExists)

    sClassName = "Function Objects"
    
    mlNumFilter1 = 11
    ReDim mobjFilter1(mlNumFilter1) As Filter1
    For i = 0 To mlNumFilter1
        Set mobjFilter1(i) = New Filter1
    Next i
    
    mlNumRampTowards = 68
    ReDim mobjRampTowards(mlNumRampTowards) As Ramp1
    For i = 0 To mlNumRampTowards
        Set mobjRampTowards(i) = New Ramp1
    Next i
    
    mlNumFollow = 30
    ReDim mobjFollow(mlNumFollow) As Follow
    For i = 0 To mlNumFollow
        Set mobjFollow(i) = New Follow
    Next i
    
    mlNumFollow2 = 27
    ReDim mobjFollow2(mlNumFollow2) As Follow2
    For i = 0 To mlNumFollow2
        Set mobjFollow2(i) = New Follow2
    Next i
    
    On Error GoTo Error
    
' Set up loop timing statistics
    Set mobjStats = New AdvUtilLib.Stats
    Set tmrElapsed = New ccrpStopWatch
    
    DoEvents
    
' Set up interlock objects
    With mobjInterlocks.Add(CStr(ocInitialize))
        .Expressions.AddBoolean "Initialized", mbInitialized, True, ilckVariableEqualValueMeansOK, LoadResString(stsControlSystemNotInitialized)
    End With

'
'' Show Status interlock
'    With objInterlocks.Add(CStr(ocShowStatus))
'        .IncludedInterlocks.Add CStr(ocStartFan)
'        .IncludedInterlocks.Add CStr(ocStop)
'    End With
    
' Set up invalid opcode collections. Add arrays of invalid opcodes to the collection and use
' the opcode as a string as an index
    With mcolInvalidOpcodes
'        .Add Array(ocShowStatus), CStr(ocShowStatus)
    End With
                    
    On Error Resume Next
    Set mobjGlobalVars = New GlobalVars
    With mobjGlobalVars
        .AddBoolean "Initialized", mbInitialized
        .AddSingle "mgvfListParsMissing 1=True", mgvfListParsMissing

' >>> Beginning of [Global Variables] .AddSingle - do not move or change this line
        .AddSingle "AT_Previous", mgvfAT_Previous
        .AddSingle "AT_Sim", mgvfAT_Sim
        .AddSingle "AT_Sim_002", mgvfAT_Sim_002
        .AddSingle "AT_Sim_003", mgvfAT_Sim_003
        .AddSingle "AT_Sim_004", mgvfAT_Sim_004
        .AddSingle "AT_Sim_005", mgvfAT_Sim_005
        .AddSingle "AT_Sim_006", mgvfAT_Sim_006
        .AddSingle "CondensTemp", mgvfCondensTemp
        .AddSingle "DP_Sim", mgvfDP_Sim
        .AddSingle "DP_Sim_001", mgvfDP_Sim_001
        .AddSingle "DP_Sim_002", mgvfDP_Sim_002
        .AddSingle "DV_Sim", mgvfDV_Sim
        .AddSingle "DY_Sim", mgvfDY_Sim
        .AddSingle "EX_Sim", mgvfEX_Sim
        .AddSingle "FaultsSim", mgvfFaultsSim
        .AddSingle "HD_Sim", mgvfHD_Sim
        .AddSingle "HD_Sim_001", mgvfHD_Sim_001
        .AddSingle "HD_Sim_002", mgvfHD_Sim_002
        .AddSingle "HD_Sim_003", mgvfHD_Sim_003
        .AddSingle "HR_Sim", mgvfHR_Sim
        .AddSingle "Hx_AvgTemp", mgvfHx_AvgTemp
        .AddSingle "Idle_Sim", mgvfIdle_Sim
        .AddSingle "MF_Previous", mgvfMF_Previous
        .AddSingle "MF_Sim_001", mgvfMF_Sim_001
        .AddSingle "MF_Sim_002", mgvfMF_Sim_002
        .AddSingle "MF_SpeedDelayed", mgvfMF_SpeedDelayed
        .AddSingle "MF_SpeedDelayType", mgvfMF_SpeedDelayType
        .AddSingle "PG_Sim", mgvfPG_Sim
        .AddSingle "Pstatic", mgvfPstatic
        .AddSingle "Q_Sim", mgvfQ_Sim
        .AddSingle "Q_Sim_001", mgvfQ_Sim_001
        .AddSingle "ResetSim", mgvfResetSim
        .AddSingle "SF_Sim", mgvfSF_Sim
        .AddSingle "SR_Sim", mgvfSR_Sim
        .AddSingle "SS_Sim", mgvfSS_Sim
        .AddSingle "TimeInterval_Sim", mgvfTimeInterval_Sim
        .AddSingle "TimeInterval_Sim_001", mgvfTimeInterval_Sim_001
        .AddSingle "WS_Sim", mgvfWS_Sim
        .AddSingle "WS_Sim_001", mgvfWS_Sim_001
' <<< End of [Global Variables] .AddSingle - do not move or change this line
    
        ' Within this simulator controller, outputs are constantly written and so cannot be changed
        ' in a GUI table (using SetParameter <ParName>, <value>)
        ' The following are copied from "Output Parameters"
        ' and copied to Global Variables so that you can use PCLook (global variables tab)
        ' to modify outputs which are only initialized once.
        ' This is useful to simulate error conditions, etc.
        
' >>> Beginning of [Global Variables] Output Parameters .AddSingle - do not move or change this line
        .AddSingle "Airtemp", mopAirtemp
        .AddSingle "AL1_FaultSts", mopAL1_FaultSts
        .AddSingle "AL2_FaultSts", mopAL2_FaultSts
        .AddSingle "BP_Raw", mopBP_Raw
        .AddSingle "CCE_Perm", mopCCE_Perm
        .AddSingle "CCE_Remote", mopCCE_Remote
        .AddSingle "CCE_RunSts", mopCCE_RunSts
        .AddSingle "CS_BrineSts", mopCS_BrineSts
        .AddSingle "CS_ChSumFlt", mopCS_ChSumFlt
        .AddSingle "CS_CritFault", mopCS_CritFault
        .AddSingle "CS_GenFault", mopCS_GenFault
        .AddSingle "CS_Health", mopCS_Health
        .AddSingle "CS_HtrFltSts", mopCS_HtrFltSts
        .AddSingle "CS_HtrRunSts", mopCS_HtrRunSts
        .AddSingle "CS_LTclsSts", mopCS_LTclsSts
        .AddSingle "CS_LTopnSts", mopCS_LTopnSts
        .AddSingle "CS_MTclsSts", mopCS_MTclsSts
        .AddSingle "CS_MTopnSts", mopCS_MTopnSts
        .AddSingle "CS_RunSts", mopCS_RunSts
        .AddSingle "Ctrl_Power", mopCtrl_Power
        .AddSingle "CV_Position", mopCV_Position
        .AddSingle "DCA_FaultSts", mopDCA_FaultSts
        .AddSingle "DCA_Perm", mopDCA_Perm
        .AddSingle "DCA_Remote", mopDCA_Remote
        .AddSingle "DCA_RunSts", mopDCA_RunSts
        .AddSingle "DCA_Temp", mopDCA_Temp
        .AddSingle "DeltaP_Hi", mopDeltaP_Hi
        .AddSingle "DeltaP_Lo", mopDeltaP_Lo
        .AddSingle "Door_Dy", mopDoor_Dy
        .AddSingle "Door_UpLg_AS", mopDoor_UpLg_AS
        .AddSingle "DV_Gly_Temp", mopDV_Gly_Temp
        .AddSingle "DV_Oil_Temp", mopDV_Oil_Temp
        .AddSingle "Dy_Estop_Sts", mopDy_Estop_Sts
        .AddSingle "Dy_Fault", mopDy_Fault
        .AddSingle "DyPit_Temp", mopDyPit_Temp
        .AddSingle "EPB_Ctrl_Rm", mopEPB_Ctrl_Rm
        .AddSingle "EPB_TS_NW", mopEPB_TS_NW
        .AddSingle "EPB_TS_SW", mopEPB_TS_SW
        .AddSingle "EPB_UpLeg_AS", mopEPB_UpLeg_AS
        .AddSingle "EStop_SumSts", mopEStop_SumSts
        .AddSingle "EX_FanSpeed", mopEX_FanSpeed
        .AddSingle "EX_Fault", mopEX_Fault
        .AddSingle "EX_Perm", mopEX_Perm
        .AddSingle "EX_Remote", mopEX_Remote
        .AddSingle "EX_RunSts", mopEX_RunSts
        .AddSingle "FaultBus", mopFaultBus
        .AddSingle "Fire_DYSup", mopFire_DYSup
        .AddSingle "Fire_Sum", mopFire_Sum
        .AddSingle "Gas_Sys_Sts", mopGas_Sys_Sts
        .AddSingle "HD_Mirror", mopHD_Mirror
        .AddSingle "HD_SmplFlow", mopHD_SmplFlow
        .AddSingle "HR_Brn_Temp", mopHR_Brn_Temp
        .AddSingle "HR_EstopSts", mopHR_EstopSts
        .AddSingle "HR_FaultSts", mopHR_FaultSts
        .AddSingle "HR_Perm", mopHR_Perm
        .AddSingle "HR1_FaultSts", mopHR1_FaultSts
        .AddSingle "HR1_HxInTmp", mopHR1_HxInTmp
        .AddSingle "HR1_HxOutTmp", mopHR1_HxOutTmp
        .AddSingle "HR1_PnlTemp", mopHR1_PnlTemp
        .AddSingle "HR2_FaultSts", mopHR2_FaultSts
        .AddSingle "HR2_HxInTmp", mopHR2_HxInTmp
        .AddSingle "HR2_HxOutTmp", mopHR2_HxOutTmp
        .AddSingle "HR2_PnlTemp", mopHR2_PnlTemp
        .AddSingle "HR3_FaultSts", mopHR3_FaultSts
        .AddSingle "HR3_HxInTmp", mopHR3_HxInTmp
        .AddSingle "HR3_HxOutTmp", mopHR3_HxOutTmp
        .AddSingle "HR3_PnlTemp", mopHR3_PnlTemp
        .AddSingle "HR4_FaultSts", mopHR4_FaultSts
        .AddSingle "HR4_HxInTmp", mopHR4_HxInTmp
        .AddSingle "HR4_HxOutTmp", mopHR4_HxOutTmp
        .AddSingle "HR4_PnlTemp", mopHR4_PnlTemp
        .AddSingle "HX_HiDeltaP", mopHX_HiDeltaP
        .AddSingle "HxIn_Temp", mopHxIn_Temp
        .AddSingle "HxOut_Temp", mopHxOut_Temp
        .AddSingle "IDL_Fault", mopIDL_Fault
        .AddSingle "IDL_FlCldSts", mopIDL_FlCldSts
        .AddSingle "IDL_FlOpdSts", mopIDL_FlOpdSts
        .AddSingle "IDL_ModeOff", mopIDL_ModeOff
        .AddSingle "IDL_ModeOn", mopIDL_ModeOn
        .AddSingle "IDL_Perm", mopIDL_Perm
        .AddSingle "IDL_Remote", mopIDL_Remote
        .AddSingle "IDL_ShCldSts", mopIDL_ShCldSts
        .AddSingle "IDL_ShOpdSts", mopIDL_ShOpdSts
        .AddSingle "Key_Dy_Door", mopKey_Dy_Door
        .AddSingle "Key_Veh_Door", mopKey_Veh_Door
        .AddSingle "Lo_Plen_P", mopLo_Plen_P
        .AddSingle "LoLo_Plen_P", mopLoLo_Plen_P
        .AddSingle "LTBrn_Temp", mopLTBrn_Temp
        .AddSingle "MF_Fault", mopMF_Fault
        .AddSingle "MF_Perm", mopMF_Perm
        .AddSingle "MF_Remote", mopMF_Remote
        .AddSingle "MF_RunSts", mopMF_RunSts
        .AddSingle "MF_SpdLim1", mopMF_SpdLim1
        .AddSingle "MF_SpdLim2", mopMF_SpdLim2
        .AddSingle "MF_Speed", mopMF_Speed
        .AddSingle "MF_SupplySts", mopMF_SupplySts
        .AddSingle "MF_Warning", mopMF_Warning
        .AddSingle "MixBrn_Temp", mopMixBrn_Temp
        .AddSingle "MTBrn_Temp", mopMTBrn_Temp
        .AddSingle "MUA_DP", mopMUA_DP
        .AddSingle "MUA_Flow", mopMUA_Flow
        .AddSingle "MUA_HX109Flt", mopMUA_HX109Flt
        .AddSingle "MUA_Temp", mopMUA_Temp
        .AddSingle "P201_FltSts", mopP201_FltSts
        .AddSingle "P201_RemSts", mopP201_RemSts
        .AddSingle "P201_RunSts", mopP201_RunSts
        .AddSingle "Pabs", mopPabs
        .AddSingle "PG_FanFault", mopPG_FanFault
        .AddSingle "PG_FanPerm", mopPG_FanPerm
        .AddSingle "PG_FanRemote", mopPG_FanRemote
        .AddSingle "PG_FanRunH", mopPG_FanRunH
        .AddSingle "PG_FanRunL", mopPG_FanRunL
        .AddSingle "PG_HL_Tmr", mopPG_HL_Tmr
        .AddSingle "PG_HLActive", mopPG_HLActive
        .AddSingle "PG_HLRunSts", mopPG_HLRunSts
        .AddSingle "PGDy_FanFlt", mopPGDy_FanFlt
        .AddSingle "PGDy_FanPerm", mopPGDy_FanPerm
        .AddSingle "PGDy_FRunH", mopPGDy_FRunH
        .AddSingle "PGDy_FRunL", mopPGDy_FRunL
        .AddSingle "PGDy_HL_Tmr", mopPGDy_HL_Tmr
        .AddSingle "PGDy_HLRnSts", mopPGDy_HLRnSts
        .AddSingle "PGDy_Remote", mopPGDy_Remote
        .AddSingle "RelHumidity", mopRelHumidity
        .AddSingle "RV_Position", mopRV_Position
        .AddSingle "SR1_AirTemp", mopSR1_AirTemp
        .AddSingle "SR1_FaultSts", mopSR1_FaultSts
        .AddSingle "SR1_Master", mopSR1_Master
        .AddSingle "SR1_RunSts", mopSR1_RunSts
        .AddSingle "SR1_SetptFbk", mopSR1_SetptFbk
        .AddSingle "SR2_AirTemp", mopSR2_AirTemp
        .AddSingle "SR2_FaultSts", mopSR2_FaultSts
        .AddSingle "SR2_Master", mopSR2_Master
        .AddSingle "SR2_RunSts", mopSR2_RunSts
        .AddSingle "SR2_SetptFbk", mopSR2_SetptFbk
        .AddSingle "SR3_AirTemp", mopSR3_AirTemp
        .AddSingle "SR3_FaultSts", mopSR3_FaultSts
        .AddSingle "SR3_Master", mopSR3_Master
        .AddSingle "SR3_RunSts", mopSR3_RunSts
        .AddSingle "SR3_SetptFbk", mopSR3_SetptFbk
        .AddSingle "SR4_AirTemp", mopSR4_AirTemp
        .AddSingle "SR4_FaultSts", mopSR4_FaultSts
        .AddSingle "SR4_Master", mopSR4_Master
        .AddSingle "SR4_RunSts", mopSR4_RunSts
        .AddSingle "SR4_SetptFbk", mopSR4_SetptFbk
        .AddSingle "SR5_AirTemp", mopSR5_AirTemp
        .AddSingle "SR5_FaultSts", mopSR5_FaultSts
        .AddSingle "SR5_Master", mopSR5_Master
        .AddSingle "SR5_RunSts", mopSR5_RunSts
        .AddSingle "SR5_SetptFbk", mopSR5_SetptFbk
        .AddSingle "SR6_AirTemp", mopSR6_AirTemp
        .AddSingle "SR6_FaultSts", mopSR6_FaultSts
        .AddSingle "SR6_Master", mopSR6_Master
        .AddSingle "SR6_RunSts", mopSR6_RunSts
        .AddSingle "SR6_SetptFbk", mopSR6_SetptFbk
        .AddSingle "SS_ArmExtSts", mopSS_ArmExtSts
        .AddSingle "SS_ArmRetSts", mopSS_ArmRetSts
        .AddSingle "SS_Fan1Fault", mopSS_Fan1Fault
        .AddSingle "SS_Fan1Rem", mopSS_Fan1Rem
        .AddSingle "SS_Fan2Fault", mopSS_Fan2Fault
        .AddSingle "SS_Fan2Rem", mopSS_Fan2Rem
        .AddSingle "SS_FanPerm", mopSS_FanPerm
        .AddSingle "SS_Fn1RunSts", mopSS_Fn1RunSts
        .AddSingle "SS_Fn2RunSts", mopSS_Fn2RunSts
        .AddSingle "SS_HiTemp", mopSS_HiTemp
        .AddSingle "SS_InputPyro", mopSS_InputPyro
        .AddSingle "SS_Perm", mopSS_Perm
        .AddSingle "SS_RunSts", mopSS_RunSts
        .AddSingle "SteamPress", mopSteamPress
        .AddSingle "SV_Position", mopSV_Position
        .AddSingle "TA1_FaultSts", mopTA1_FaultSts
        .AddSingle "TA2_FaultSts", mopTA2_FaultSts
        .AddSingle "VH_IgnPerm", mopVH_IgnPerm
        .AddSingle "VH_IgnSts", mopVH_IgnSts
' <<< End of [Global Variables] Output Parameters .AddSingle - do not move or change this line

    End With
    
    On Error GoTo Error
    
' Set up Statuses
    Set mobjStatuses = New Statuses
    With mobjStatuses
        .UseSPG SPG_STATUSES
        .Add STS_SIM_INITIALIZED, False, "Sim Not Initialized", "Sim Initialized"
    End With
    
' Create a process object, and then pass this Process
' object to the Process Manager

    Set objProcess = New Process
    mbReady = True
    
    WaitForProcessesStartupComplete
    
'    Set objProcessManagerProcess = AddTalentProcessWait(CStr(objProcess.Name), objProcess)
'    lID = objProcessManagerProcess.ID
    mlID = AddTalentProcessGetID(CStr(objProcess.Name), objProcess)
            
' Load control constants
    On Error Resume Next
'    Initialize initConstants
    On Error GoTo Error
    
' Set up parameter objects (after adding process and getting Process ID)

    On Error GoTo Error_Configurations
    sClassName = "ParManager"
    Set mobjParManager = New ParManager
    mobjParManager.OwnerID = mlID
    With mobjParManager
'        .AutoRefreshEnabled = False
        .AutoRefreshEnabled = True
        .IgnoreLoadedEvents = True
        .IgnoreUnloadedEvents = True

    ' Input Parameters
' >>> Beginning of [Input Parameters] .AddWithoutFlag - do not move or change this line
        .AddWithoutFlag "AirDensity", mipAirDensity, ioInput
        .AddWithoutFlag "AT_Rate", mipAT_Rate, ioInput
        .AddWithoutFlag "CCE_Start", mipCCE_Start, ioInput
        .AddWithoutFlag "CS_BrineReq", mipCS_BrineReq, ioInput
        .AddWithoutFlag "CS_HtrSetpt", mipCS_HtrSetpt, ioInput
        .AddWithoutFlag "CV_Setpt", mipCV_Setpt, ioInput
        .AddWithoutFlag "DCA_Start", mipDCA_Start, ioInput
        .AddWithoutFlag "DCA_Target", mipDCA_Target, ioInput
        .AddWithoutFlag "Dewpoint", mipDewpoint, ioInput
        .AddWithoutFlag "DP_Target", mipDP_Target, ioInput
        .AddWithoutFlag "DV_Gly_Start", mipDV_Gly_Start, ioInput
        .AddWithoutFlag "DV_Gly_TTgt", mipDV_Gly_TTgt, ioInput
        .AddWithoutFlag "DV_Oil_Start", mipDV_Oil_Start, ioInput
        .AddWithoutFlag "DV_Oil_TTgt", mipDV_Oil_TTgt, ioInput
        .AddWithoutFlag "Dy_CommSts", mipDy_CommSts, ioInput
        .AddWithoutFlag "Dy_Remote", mipDy_Remote, ioInput
        .AddWithoutFlag "DyPit_Start", mipDyPit_Start, ioInput
        .AddWithoutFlag "DyPit_Target", mipDyPit_Target, ioInput
        .AddWithoutFlag "EX_FanSetpt", mipEX_FanSetpt, ioInput
        .AddWithoutFlag "EX_Start", mipEX_Start, ioInput
        .AddWithoutFlag "HR_BrnSetpt", mipHR_BrnSetpt, ioInput
        .AddWithoutFlag "HR_RunSts", mipHR_RunSts, ioInput
        .AddWithoutFlag "HR1_CVSetpt", mipHR1_CVSetpt, ioInput
        .AddWithoutFlag "HR1_HtrSetpt", mipHR1_HtrSetpt, ioInput
        .AddWithoutFlag "HR1_RunSts", mipHR1_RunSts, ioInput
        .AddWithoutFlag "HR2_CVSetpt", mipHR2_CVSetpt, ioInput
        .AddWithoutFlag "HR2_HtrSetpt", mipHR2_HtrSetpt, ioInput
        .AddWithoutFlag "HR2_RunSts", mipHR2_RunSts, ioInput
        .AddWithoutFlag "HR3_CVSetpt", mipHR3_CVSetpt, ioInput
        .AddWithoutFlag "HR3_HtrSetpt", mipHR3_HtrSetpt, ioInput
        .AddWithoutFlag "HR3_RunSts", mipHR3_RunSts, ioInput
        .AddWithoutFlag "HR4_CVSetpt", mipHR4_CVSetpt, ioInput
        .AddWithoutFlag "HR4_HtrSetpt", mipHR4_HtrSetpt, ioInput
        .AddWithoutFlag "HR4_RunSts", mipHR4_RunSts, ioInput
        .AddWithoutFlag "IDL_Start", mipIDL_Start, ioInput
        .AddWithoutFlag "LTBrn_Warn", mipLTBrn_Warn, ioInput
        .AddWithoutFlag "MF_Setpt", mipMF_Setpt, ioInput
        .AddWithoutFlag "MF_Start", mipMF_Start, ioInput
        .AddWithoutFlag "MTBrn_Warn", mipMTBrn_Warn, ioInput
        .AddWithoutFlag "P201_RunReq", mipP201_RunReq, ioInput
        .AddWithoutFlag "PG_HL_DelRst", mipPG_HL_DelRst, ioInput
        .AddWithoutFlag "PG_HL_Start", mipPG_HL_Start, ioInput
        .AddWithoutFlag "PG_LL_Start", mipPG_LL_Start, ioInput
        .AddWithoutFlag "PGDy_HL_DRst", mipPGDy_HL_DRst, ioInput
        .AddWithoutFlag "PGDy_HL_Strt", mipPGDy_HL_Strt, ioInput
        .AddWithoutFlag "PGDy_LL_Strt", mipPGDy_LL_Strt, ioInput
        .AddWithoutFlag "PStatic", mipPStatic, ioInput
        .AddWithoutFlag "Q", mipQ, ioInput
        .AddWithoutFlag "RV_Setpt", mipRV_Setpt, ioInput
        .AddWithoutFlag "SR1_RunReq", mipSR1_RunReq, ioInput
        .AddWithoutFlag "SR1_Setpt", mipSR1_Setpt, ioInput
        .AddWithoutFlag "SR2_RunReq", mipSR2_RunReq, ioInput
        .AddWithoutFlag "SR2_Setpt", mipSR2_Setpt, ioInput
        .AddWithoutFlag "SR3_RunReq", mipSR3_RunReq, ioInput
        .AddWithoutFlag "SR3_Setpt", mipSR3_Setpt, ioInput
        .AddWithoutFlag "SR4_RunReq", mipSR4_RunReq, ioInput
        .AddWithoutFlag "SR4_Setpt", mipSR4_Setpt, ioInput
        .AddWithoutFlag "SR5_RunReq", mipSR5_RunReq, ioInput
        .AddWithoutFlag "SR5_Setpt", mipSR5_Setpt, ioInput
        .AddWithoutFlag "SR6_RunReq", mipSR6_RunReq, ioInput
        .AddWithoutFlag "SR6_Setpt", mipSR6_Setpt, ioInput
        .AddWithoutFlag "SS_ArmExtReq", mipSS_ArmExtReq, ioInput
        .AddWithoutFlag "SS_FanStart", mipSS_FanStart, ioInput
        .AddWithoutFlag "SS_InvCoef0", mipSS_InvCoef0, ioInput
        .AddWithoutFlag "SS_InvCoef1", mipSS_InvCoef1, ioInput
        .AddWithoutFlag "SS_InvCoef2", mipSS_InvCoef2, ioInput
        .AddWithoutFlag "SS_InvCoef3", mipSS_InvCoef3, ioInput
        .AddWithoutFlag "SS_InvCoef4", mipSS_InvCoef4, ioInput
        .AddWithoutFlag "SS_Power", mipSS_Power, ioInput
        .AddWithoutFlag "SS_RunReq", mipSS_RunReq, ioInput
        .AddWithoutFlag "SV_Setpt", mipSV_Setpt, ioInput
        .AddWithoutFlag "VH_IgnEnable", mipVH_IgnEnable, ioInput
        .AddWithoutFlag "WS_Setpt", mipWS_Setpt, ioInput
' <<< End of [Input Parameters] .AddWithoutFlag - do not move or change this line

'        For i = 1 To 30
'            .AddWithoutFlag "RainValve" & Format(i, "00"), mipRainValve(i), ioInput
'        Next i
    
    ' Output Parameters
' >>> Beginning of [Output Parameters] .AddWithoutFlag - do not move or change this line
        .AddWithoutFlag "Airtemp", mopAirtemp, ioOutput
        .AddWithoutFlag "AL1_FaultSts", mopAL1_FaultSts, ioOutput
        .AddWithoutFlag "AL2_FaultSts", mopAL2_FaultSts, ioOutput
        .AddWithoutFlag "BP_Raw", mopBP_Raw, ioOutput
        .AddWithoutFlag "CCE_Perm", mopCCE_Perm, ioOutput
        .AddWithoutFlag "CCE_Remote", mopCCE_Remote, ioOutput
        .AddWithoutFlag "CCE_RunSts", mopCCE_RunSts, ioOutput
        .AddWithoutFlag "CS_BrineSts", mopCS_BrineSts, ioOutput
        .AddWithoutFlag "CS_ChSumFlt", mopCS_ChSumFlt, ioOutput
        .AddWithoutFlag "CS_CritFault", mopCS_CritFault, ioOutput
        .AddWithoutFlag "CS_GenFault", mopCS_GenFault, ioOutput
        .AddWithoutFlag "CS_Health", mopCS_Health, ioOutput
        .AddWithoutFlag "CS_HtrFltSts", mopCS_HtrFltSts, ioOutput
        .AddWithoutFlag "CS_HtrRunSts", mopCS_HtrRunSts, ioOutput
        .AddWithoutFlag "CS_LTclsSts", mopCS_LTclsSts, ioOutput
        .AddWithoutFlag "CS_LTopnSts", mopCS_LTopnSts, ioOutput
        .AddWithoutFlag "CS_MTclsSts", mopCS_MTclsSts, ioOutput
        .AddWithoutFlag "CS_MTopnSts", mopCS_MTopnSts, ioOutput
        .AddWithoutFlag "CS_RunSts", mopCS_RunSts, ioOutput
        .AddWithoutFlag "Ctrl_Power", mopCtrl_Power, ioOutput
        .AddWithoutFlag "CV_Position", mopCV_Position, ioOutput
        .AddWithoutFlag "DCA_FaultSts", mopDCA_FaultSts, ioOutput
        .AddWithoutFlag "DCA_Perm", mopDCA_Perm, ioOutput
        .AddWithoutFlag "DCA_Remote", mopDCA_Remote, ioOutput
        .AddWithoutFlag "DCA_RunSts", mopDCA_RunSts, ioOutput
        .AddWithoutFlag "DCA_Temp", mopDCA_Temp, ioOutput
        .AddWithoutFlag "DeltaP_Hi", mopDeltaP_Hi, ioOutput
        .AddWithoutFlag "DeltaP_Lo", mopDeltaP_Lo, ioOutput
        .AddWithoutFlag "Door_Dy", mopDoor_Dy, ioOutput
        .AddWithoutFlag "Door_UpLg_AS", mopDoor_UpLg_AS, ioOutput
        .AddWithoutFlag "DV_Gly_Temp", mopDV_Gly_Temp, ioOutput
        .AddWithoutFlag "DV_Oil_Temp", mopDV_Oil_Temp, ioOutput
        .AddWithoutFlag "Dy_Estop_Sts", mopDy_Estop_Sts, ioOutput
        .AddWithoutFlag "Dy_Fault", mopDy_Fault, ioOutput
        .AddWithoutFlag "DyPit_Temp", mopDyPit_Temp, ioOutput
        .AddWithoutFlag "EPB_Ctrl_Rm", mopEPB_Ctrl_Rm, ioOutput
        .AddWithoutFlag "EPB_TS_NW", mopEPB_TS_NW, ioOutput
        .AddWithoutFlag "EPB_TS_SW", mopEPB_TS_SW, ioOutput
        .AddWithoutFlag "EPB_UpLeg_AS", mopEPB_UpLeg_AS, ioOutput
        .AddWithoutFlag "EStop_SumSts", mopEStop_SumSts, ioOutput
        .AddWithoutFlag "EX_FanSpeed", mopEX_FanSpeed, ioOutput
        .AddWithoutFlag "EX_Fault", mopEX_Fault, ioOutput
        .AddWithoutFlag "EX_Perm", mopEX_Perm, ioOutput
        .AddWithoutFlag "EX_Remote", mopEX_Remote, ioOutput
        .AddWithoutFlag "EX_RunSts", mopEX_RunSts, ioOutput
        .AddWithoutFlag "FaultBus", mopFaultBus, ioOutput
        .AddWithoutFlag "Fire_DYSup", mopFire_DYSup, ioOutput
        .AddWithoutFlag "Fire_Sum", mopFire_Sum, ioOutput
        .AddWithoutFlag "Gas_Sys_Sts", mopGas_Sys_Sts, ioOutput
        .AddWithoutFlag "HD_Mirror", mopHD_Mirror, ioOutput
        .AddWithoutFlag "HD_SmplFlow", mopHD_SmplFlow, ioOutput
        .AddWithoutFlag "HR_Brn_Temp", mopHR_Brn_Temp, ioOutput
        .AddWithoutFlag "HR_EstopSts", mopHR_EstopSts, ioOutput
        .AddWithoutFlag "HR_FaultSts", mopHR_FaultSts, ioOutput
        .AddWithoutFlag "HR_Perm", mopHR_Perm, ioOutput
        .AddWithoutFlag "HR1_FaultSts", mopHR1_FaultSts, ioOutput
        .AddWithoutFlag "HR1_HxInTmp", mopHR1_HxInTmp, ioOutput
        .AddWithoutFlag "HR1_HxOutTmp", mopHR1_HxOutTmp, ioOutput
        .AddWithoutFlag "HR1_PnlTemp", mopHR1_PnlTemp, ioOutput
        .AddWithoutFlag "HR2_FaultSts", mopHR2_FaultSts, ioOutput
        .AddWithoutFlag "HR2_HxInTmp", mopHR2_HxInTmp, ioOutput
        .AddWithoutFlag "HR2_HxOutTmp", mopHR2_HxOutTmp, ioOutput
        .AddWithoutFlag "HR2_PnlTemp", mopHR2_PnlTemp, ioOutput
        .AddWithoutFlag "HR3_FaultSts", mopHR3_FaultSts, ioOutput
        .AddWithoutFlag "HR3_HxInTmp", mopHR3_HxInTmp, ioOutput
        .AddWithoutFlag "HR3_HxOutTmp", mopHR3_HxOutTmp, ioOutput
        .AddWithoutFlag "HR3_PnlTemp", mopHR3_PnlTemp, ioOutput
        .AddWithoutFlag "HR4_FaultSts", mopHR4_FaultSts, ioOutput
        .AddWithoutFlag "HR4_HxInTmp", mopHR4_HxInTmp, ioOutput
        .AddWithoutFlag "HR4_HxOutTmp", mopHR4_HxOutTmp, ioOutput
        .AddWithoutFlag "HR4_PnlTemp", mopHR4_PnlTemp, ioOutput
        .AddWithoutFlag "HX_HiDeltaP", mopHX_HiDeltaP, ioOutput
        .AddWithoutFlag "HxIn_Temp", mopHxIn_Temp, ioOutput
        .AddWithoutFlag "HxOut_Temp", mopHxOut_Temp, ioOutput
        .AddWithoutFlag "IDL_Fault", mopIDL_Fault, ioOutput
        .AddWithoutFlag "IDL_FlCldSts", mopIDL_FlCldSts, ioOutput
        .AddWithoutFlag "IDL_FlOpdSts", mopIDL_FlOpdSts, ioOutput
        .AddWithoutFlag "IDL_ModeOff", mopIDL_ModeOff, ioOutput
        .AddWithoutFlag "IDL_ModeOn", mopIDL_ModeOn, ioOutput
        .AddWithoutFlag "IDL_Perm", mopIDL_Perm, ioOutput
        .AddWithoutFlag "IDL_Remote", mopIDL_Remote, ioOutput
        .AddWithoutFlag "IDL_ShCldSts", mopIDL_ShCldSts, ioOutput
        .AddWithoutFlag "IDL_ShOpdSts", mopIDL_ShOpdSts, ioOutput
        .AddWithoutFlag "Key_Dy_Door", mopKey_Dy_Door, ioOutput
        .AddWithoutFlag "Key_Veh_Door", mopKey_Veh_Door, ioOutput
        .AddWithoutFlag "Lo_Plen_P", mopLo_Plen_P, ioOutput
        .AddWithoutFlag "LoLo_Plen_P", mopLoLo_Plen_P, ioOutput
        .AddWithoutFlag "LTBrn_Temp", mopLTBrn_Temp, ioOutput
        .AddWithoutFlag "MF_Fault", mopMF_Fault, ioOutput
        .AddWithoutFlag "MF_Perm", mopMF_Perm, ioOutput
        .AddWithoutFlag "MF_Remote", mopMF_Remote, ioOutput
        .AddWithoutFlag "MF_RunSts", mopMF_RunSts, ioOutput
        .AddWithoutFlag "MF_SpdLim1", mopMF_SpdLim1, ioOutput
        .AddWithoutFlag "MF_SpdLim2", mopMF_SpdLim2, ioOutput
        .AddWithoutFlag "MF_Speed", mopMF_Speed, ioOutput
        .AddWithoutFlag "MF_SupplySts", mopMF_SupplySts, ioOutput
        .AddWithoutFlag "MF_Warning", mopMF_Warning, ioOutput
        .AddWithoutFlag "MixBrn_Temp", mopMixBrn_Temp, ioOutput
        .AddWithoutFlag "MTBrn_Temp", mopMTBrn_Temp, ioOutput
        .AddWithoutFlag "MUA_DP", mopMUA_DP, ioOutput
        .AddWithoutFlag "MUA_Flow", mopMUA_Flow, ioOutput
        .AddWithoutFlag "MUA_HX109Flt", mopMUA_HX109Flt, ioOutput
        .AddWithoutFlag "MUA_Temp", mopMUA_Temp, ioOutput
        .AddWithoutFlag "P201_FltSts", mopP201_FltSts, ioOutput
        .AddWithoutFlag "P201_RemSts", mopP201_RemSts, ioOutput
        .AddWithoutFlag "P201_RunSts", mopP201_RunSts, ioOutput
        .AddWithoutFlag "Pabs", mopPabs, ioOutput
        .AddWithoutFlag "PG_FanFault", mopPG_FanFault, ioOutput
        .AddWithoutFlag "PG_FanPerm", mopPG_FanPerm, ioOutput
        .AddWithoutFlag "PG_FanRemote", mopPG_FanRemote, ioOutput
        .AddWithoutFlag "PG_FanRunH", mopPG_FanRunH, ioOutput
        .AddWithoutFlag "PG_FanRunL", mopPG_FanRunL, ioOutput
        .AddWithoutFlag "PG_HL_Tmr", mopPG_HL_Tmr, ioOutput
        .AddWithoutFlag "PG_HLActive", mopPG_HLActive, ioOutput
        .AddWithoutFlag "PG_HLRunSts", mopPG_HLRunSts, ioOutput
        .AddWithoutFlag "PGDy_FanFlt", mopPGDy_FanFlt, ioOutput
        .AddWithoutFlag "PGDy_FanPerm", mopPGDy_FanPerm, ioOutput
        .AddWithoutFlag "PGDy_FRunH", mopPGDy_FRunH, ioOutput
        .AddWithoutFlag "PGDy_FRunL", mopPGDy_FRunL, ioOutput
        .AddWithoutFlag "PGDy_HL_Tmr", mopPGDy_HL_Tmr, ioOutput
        .AddWithoutFlag "PGDy_HLRnSts", mopPGDy_HLRnSts, ioOutput
        .AddWithoutFlag "PGDy_Remote", mopPGDy_Remote, ioOutput
        .AddWithoutFlag "RelHumidity", mopRelHumidity, ioOutput
        .AddWithoutFlag "RV_Position", mopRV_Position, ioOutput
        .AddWithoutFlag "SR1_AirTemp", mopSR1_AirTemp, ioOutput
        .AddWithoutFlag "SR1_FaultSts", mopSR1_FaultSts, ioOutput
        .AddWithoutFlag "SR1_Master", mopSR1_Master, ioOutput
        .AddWithoutFlag "SR1_RunSts", mopSR1_RunSts, ioOutput
        .AddWithoutFlag "SR1_SetptFbk", mopSR1_SetptFbk, ioOutput
        .AddWithoutFlag "SR2_AirTemp", mopSR2_AirTemp, ioOutput
        .AddWithoutFlag "SR2_FaultSts", mopSR2_FaultSts, ioOutput
        .AddWithoutFlag "SR2_Master", mopSR2_Master, ioOutput
        .AddWithoutFlag "SR2_RunSts", mopSR2_RunSts, ioOutput
        .AddWithoutFlag "SR2_SetptFbk", mopSR2_SetptFbk, ioOutput
        .AddWithoutFlag "SR3_AirTemp", mopSR3_AirTemp, ioOutput
        .AddWithoutFlag "SR3_FaultSts", mopSR3_FaultSts, ioOutput
        .AddWithoutFlag "SR3_Master", mopSR3_Master, ioOutput
        .AddWithoutFlag "SR3_RunSts", mopSR3_RunSts, ioOutput
        .AddWithoutFlag "SR3_SetptFbk", mopSR3_SetptFbk, ioOutput
        .AddWithoutFlag "SR4_AirTemp", mopSR4_AirTemp, ioOutput
        .AddWithoutFlag "SR4_FaultSts", mopSR4_FaultSts, ioOutput
        .AddWithoutFlag "SR4_Master", mopSR4_Master, ioOutput
        .AddWithoutFlag "SR4_RunSts", mopSR4_RunSts, ioOutput
        .AddWithoutFlag "SR4_SetptFbk", mopSR4_SetptFbk, ioOutput
        .AddWithoutFlag "SR5_AirTemp", mopSR5_AirTemp, ioOutput
        .AddWithoutFlag "SR5_FaultSts", mopSR5_FaultSts, ioOutput
        .AddWithoutFlag "SR5_Master", mopSR5_Master, ioOutput
        .AddWithoutFlag "SR5_RunSts", mopSR5_RunSts, ioOutput
        .AddWithoutFlag "SR5_SetptFbk", mopSR5_SetptFbk, ioOutput
        .AddWithoutFlag "SR6_AirTemp", mopSR6_AirTemp, ioOutput
        .AddWithoutFlag "SR6_FaultSts", mopSR6_FaultSts, ioOutput
        .AddWithoutFlag "SR6_Master", mopSR6_Master, ioOutput
        .AddWithoutFlag "SR6_RunSts", mopSR6_RunSts, ioOutput
        .AddWithoutFlag "SR6_SetptFbk", mopSR6_SetptFbk, ioOutput
        .AddWithoutFlag "SS_ArmExtSts", mopSS_ArmExtSts, ioOutput
        .AddWithoutFlag "SS_ArmRetSts", mopSS_ArmRetSts, ioOutput
        .AddWithoutFlag "SS_Fan1Fault", mopSS_Fan1Fault, ioOutput
        .AddWithoutFlag "SS_Fan1Rem", mopSS_Fan1Rem, ioOutput
        .AddWithoutFlag "SS_Fan2Fault", mopSS_Fan2Fault, ioOutput
        .AddWithoutFlag "SS_Fan2Rem", mopSS_Fan2Rem, ioOutput
        .AddWithoutFlag "SS_FanPerm", mopSS_FanPerm, ioOutput
        .AddWithoutFlag "SS_Fn1RunSts", mopSS_Fn1RunSts, ioOutput
        .AddWithoutFlag "SS_Fn2RunSts", mopSS_Fn2RunSts, ioOutput
        .AddWithoutFlag "SS_HiTemp", mopSS_HiTemp, ioOutput
        .AddWithoutFlag "SS_InputPyro", mopSS_InputPyro, ioOutput
        .AddWithoutFlag "SS_Perm", mopSS_Perm, ioOutput
        .AddWithoutFlag "SS_RunSts", mopSS_RunSts, ioOutput
        .AddWithoutFlag "SteamPress", mopSteamPress, ioOutput
        .AddWithoutFlag "SV_Position", mopSV_Position, ioOutput
        .AddWithoutFlag "TA1_FaultSts", mopTA1_FaultSts, ioOutput
        .AddWithoutFlag "TA2_FaultSts", mopTA2_FaultSts, ioOutput
        .AddWithoutFlag "VH_IgnPerm", mopVH_IgnPerm, ioOutput
        .AddWithoutFlag "VH_IgnSts", mopVH_IgnSts, ioOutput
' <<< End of [Output Parameters] .AddWithoutFlag - do not move or change this line

'        For i = 1 To 28
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "Control", mudtSSLamp(i).Control, ioInput
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "Status", mudtSSLamp(i).Status, ioOutput
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "Selected", mudtSSLamp(i).Selected, ioInput
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "SelectedStatus", mudtSSLamp(i).SelectedStatus, ioOutput
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "Power", mudtSSLamp(i).Power, ioOutput
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "CloudShutterCmd", mudtSSLamp(i).CloudShutterCmd, ioInput
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "CloudShutterStatus", mudtSSLamp(i).CloudShutterStatus, ioOutput
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "TunnelShutterCmd", mudtSSLamp(i).TunnelShutterCmd, ioInput
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "TunnelShutterStatus", mudtSSLamp(i).TunnelShutterStatus, ioOutput
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "OffsetSP", mudtSSLamp(i).OffsetSp, ioInput
'            .AddWithoutFlag "SSLamp" & Format(i, "00") & "Offset", mudtSSLamp(i).Offset, ioOutput
'        Next i
    End With
    
    On Error GoTo Error
    
' Now fully started up. Set process priority and enable timers
' If running within VB (the IDE) then use a form based timer, and don't
' use Real Time Priority
    If mccfUpdateTime < 0.05 Then
        mccfUpdateTime = 0.05
    End If

    If IsWithinIDE() Then
        Set mobjForm = New frmControl
        Load mobjForm
        mobjForm.tmrDesignTime.Interval = mccfUpdateTime * 1000
        mobjForm.tmrDesignTime.Enabled = True
    Else
        tmrMain.Stats.Frequency = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_PROCESSES, , "TimerFrequency", 10, vbLong)
        tmrMain.Interval = mccfUpdateTime * 1000           ' 50 milliseconds
        tmrMain.EventType = TimerPeriodic
        Set tmrMain.Notify = Me
        SetProcessPriority mlDefPriority
        tmrMain.Enabled = True
    End If

' Insert reset command to set up communications
    If mobjParManager.Configurations.Count > 0 Then
        mobjCommands.Insert ocInitialize, "SimInitialize"
    End If
    
' Clear Run Once timer
    Set tmrRunOnce = Nothing
    mobjStats.Reset
    For i = 1 To 200
        mobjStats.AddValue CDbl(mccfUpdateTime)
    Next i
    tmrElapsed.Reset

    GoTo Done
    
Error:
    sRaiseErrorText = Err.Description
'    bRaiseError = True
'    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
'    End If
    LogMessage "Error during controller startup: " & sRaiseErrorText & ". See log file for details", mtFailure
    Resume Next

Error_CreatingObjects:
    sRaiseErrorText = "Can't create object '" & sClassName & "': " & Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    EH.Item(1) = sClassName
    Call EH_MsgBox
    Resume Done
    Resume

Error_Configurations:
    Set EH.VBErr = Err
    EH.AppError = LoadResString(StdTalentStrings.stsEHCantCreateConfigurations)
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    sRaiseErrorText = EH.Text
    Call EH_MsgBox
    Resume Next
    
Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Sub

End Sub

'Private Sub AddOutput(ParameterName As String, ValueHolder As Single, ResetToDefaultUponComError As Boolean, IsCalculatedBySimulator As Boolean)
'    If mgvbSimulation Then
'        mobjParManager.AddWithoutFlag(ParameterName, ValueHolder, IIf(IsCalculatedBySimulator, ioInput, ioOutput)).Tag = ResetToDefaultUponComError
'    Else
'        mobjParManager.AddWithoutFlag(ParameterName, ValueHolder, ioOutput).Tag = ResetToDefaultUponComError
'    End If
'End Sub
Public Property Get Stats() As Stats
    Set Stats = mobjStats
End Property

Private Sub Suspend()
Attribute Suspend.VB_Description = "Suspend control"

' Suspend control

    On Error Resume Next

    mbInitialized = False
    mobjStatuses(STS_SIM_INITIALIZED).SetValue mbInitialized
    
    Exit Sub

End Sub

Private Sub tmrRunOnce_Timer(ByVal Milliseconds As Long)

'

    Const PROC_NAME = "tmrRunOnce_Timer"

    On Error GoTo Error
    
    tmrRunOnce.Enabled = False
    Set tmrRunOnce = Nothing

    Start

    GoTo Done

Error:
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next

    Exit Sub


End Sub


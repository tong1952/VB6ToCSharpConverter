// Converted from Controller.cls by vb6cs
// Date: 2026-05-11 03:23

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public class Controller : ICcrpTimerNotify
{
    private CommandListLib.Commands mobjCommands;
    private TalentProcessLib.Statuses mobjStatuses;
    private ControlLib.ControlConstants mobjConstants;
    private ControlLib.GlobalVars mobjGlobalVars;
    private AdvUtilLib.Stats mobjStats;
    private LoopTimer tmrMain;
    private ccrpTimer tmrRunOnce;
    private ccrpStopWatch tmrElapsed;
    private Interlocks mobjInterlocks;
    private TalentProcessLib.Events mobjEvents;
    private frmControl mobjForm;
    private SharedPropertyGroup mobjSPGConstants;
    private SharedPropertyGroup mobjSPGDataItems;
    private ParametersLib.ParManager mobjParManager;
    private int mlID;
    private bool mbReady;
    private bool mbInitialized;
    private const bool bKeepStats = false;
    private bool bRealTime;
    private object mvInterlockCheckValue;
    private DateTime mdStartTime;
    private const bool mbKeepStats = true;
    private bool mbResetSim;
    private float[] mfMFSpeedSPBuffer;
    private int mlMFSpeedSPBufferIndex;
    private int mlMFSpeedSPBufferSize;
    private int mlNumFilter1;
    private int mlNumRampTowards;
    private int mlNumFollow;
    private int mlNumFollow2;
    private Collection mcolInvalidOpcodes;
    private enum InitTypes
    {
        initAll = 0,
        initConstants = 1,
        initPars = 2
    }

    private string mccsPriority;
    private float mccfUpdateTime;
    private float mccfFanDelay;
    private float mccfTunnelVolume;
    private float[] mccfKp1Coeff = new float[2 + 1];
    private float[] mccfKp2Coeff = new float[2 + 1];
    private float mccfKpDpTrans;
    private float[] mccfKq1Coeff = new float[2 + 1];
    private float[] mccfKq2Coeff = new float[2 + 1];
    private float mccfKqDpTrans;
    private float mccfDPxOverHi;
    private float mccfDPxOverLo;
    private float mipAirDensity;
    private float mipAT_Rate;
    private float mipCCE_Start;
    private float mipCS_BrineReq;
    private float mipCS_HtrSetpt;
    private float mipCV_Setpt;
    private float mipDCA_Start;
    private float mipDCA_Target;
    private float mipDewpoint;
    private float mipDP_Target;
    private float mipDV_Gly_Start;
    private float mipDV_Gly_TTgt;
    private float mipDV_Oil_Start;
    private float mipDV_Oil_TTgt;
    private float mipDy_CommSts;
    private float mipDy_Remote;
    private float mipDyPit_Start;
    private float mipDyPit_Target;
    private float mipEX_FanSetpt;
    private float mipEX_Start;
    private float mipHR_BrnSetpt;
    private float mipHR_RunSts;
    private float mipHR1_CVSetpt;
    private float mipHR1_HtrSetpt;
    private float mipHR1_RunSts;
    private float mipHR2_CVSetpt;
    private float mipHR2_HtrSetpt;
    private float mipHR2_RunSts;
    private float mipHR3_CVSetpt;
    private float mipHR3_HtrSetpt;
    private float mipHR3_RunSts;
    private float mipHR4_CVSetpt;
    private float mipHR4_HtrSetpt;
    private float mipHR4_RunSts;
    private float mipIDL_Start;
    private float mipLTBrn_Warn;
    private float mipMF_Setpt;
    private float mipMF_Start;
    private float mipMTBrn_Warn;
    private float mipP201_RunReq;
    private float mipPG_HL_DelRst;
    private float mipPG_HL_Start;
    private float mipPG_LL_Start;
    private float mipPGDy_HL_DRst;
    private float mipPGDy_HL_Strt;
    private float mipPGDy_LL_Strt;
    private float mipPStatic;
    private float mipQ;
    private float mipRV_Setpt;
    private float mipSR1_RunReq;
    private float mipSR1_Setpt;
    private float mipSR2_RunReq;
    private float mipSR2_Setpt;
    private float mipSR3_RunReq;
    private float mipSR3_Setpt;
    private float mipSR4_RunReq;
    private float mipSR4_Setpt;
    private float mipSR5_RunReq;
    private float mipSR5_Setpt;
    private float mipSR6_RunReq;
    private float mipSR6_Setpt;
    private float mipSS_ArmExtReq;
    private float mipSS_FanStart;
    private float mipSS_InvCoef0;
    private float mipSS_InvCoef1;
    private float mipSS_InvCoef2;
    private float mipSS_InvCoef3;
    private float mipSS_InvCoef4;
    private float mipSS_Power;
    private float mipSS_RunReq;
    private float mipSV_Setpt;
    private float mipVH_IgnEnable;
    private float mipWS_Setpt;
    private float mopAirtemp;
    private float mopAL1_FaultSts;
    private float mopAL2_FaultSts;
    private float mopBP_Raw;
    private float mopCCE_Perm;
    private float mopCCE_Remote;
    private float mopCCE_RunSts;
    private float mopCS_BrineSts;
    private float mopCS_ChSumFlt;
    private float mopCS_CritFault;
    private float mopCS_GenFault;
    private float mopCS_Health;
    private float mopCS_HtrFltSts;
    private float mopCS_HtrRunSts;
    private float mopCS_LTclsSts;
    private float mopCS_LTopnSts;
    private float mopCS_MTclsSts;
    private float mopCS_MTopnSts;
    private float mopCS_RunSts;
    private float mopCtrl_Power;
    private float mopCV_Position;
    private float mopDCA_FaultSts;
    private float mopDCA_Perm;
    private float mopDCA_Remote;
    private float mopDCA_RunSts;
    private float mopDCA_Temp;
    private float mopDeltaP_Hi;
    private float mopDeltaP_Lo;
    private float mopDoor_Dy;
    private float mopDoor_UpLg_AS;
    private float mopDV_Gly_Temp;
    private float mopDV_Oil_Temp;
    private float mopDy_Estop_Sts;
    private float mopDy_Fault;
    private float mopDyPit_Temp;
    private float mopEPB_Ctrl_Rm;
    private float mopEPB_TS_NW;
    private float mopEPB_TS_SW;
    private float mopEPB_UpLeg_AS;
    private float mopEStop_SumSts;
    private float mopEX_FanSpeed;
    private float mopEX_Fault;
    private float mopEX_Perm;
    private float mopEX_Remote;
    private float mopEX_RunSts;
    private float mopFaultBus;
    private float mopFire_DYSup;
    private float mopFire_Sum;
    private float mopGas_Sys_Sts;
    private float mopHD_Mirror;
    private float mopHD_SmplFlow;
    private float mopHR_Brn_Temp;
    private float mopHR_EstopSts;
    private float mopHR_FaultSts;
    private float mopHR_Perm;
    private float mopHR1_FaultSts;
    private float mopHR1_HxInTmp;
    private float mopHR1_HxOutTmp;
    private float mopHR1_PnlTemp;
    private float mopHR2_FaultSts;
    private float mopHR2_HxInTmp;
    private float mopHR2_HxOutTmp;
    private float mopHR2_PnlTemp;
    private float mopHR3_FaultSts;
    private float mopHR3_HxInTmp;
    private float mopHR3_HxOutTmp;
    private float mopHR3_PnlTemp;
    private float mopHR4_FaultSts;
    private float mopHR4_HxInTmp;
    private float mopHR4_HxOutTmp;
    private float mopHR4_PnlTemp;
    private float mopHX_HiDeltaP;
    private float mopHxIn_Temp;
    private float mopHxOut_Temp;
    private float mopIDL_Fault;
    private float mopIDL_FlCldSts;
    private float mopIDL_FlOpdSts;
    private float mopIDL_ModeOff;
    private float mopIDL_ModeOn;
    private float mopIDL_Perm;
    private float mopIDL_Remote;
    private float mopIDL_ShCldSts;
    private float mopIDL_ShOpdSts;
    private float mopKey_Dy_Door;
    private float mopKey_Veh_Door;
    private float mopLo_Plen_P;
    private float mopLoLo_Plen_P;
    private float mopLTBrn_Temp;
    private float mopMF_Fault;
    private float mopMF_Perm;
    private float mopMF_Remote;
    private float mopMF_RunSts;
    private float mopMF_SpdLim1;
    private float mopMF_SpdLim2;
    private float mopMF_Speed;
    private float mopMF_SupplySts;
    private float mopMF_Warning;
    private float mopMixBrn_Temp;
    private float mopMTBrn_Temp;
    private float mopMUA_DP;
    private float mopMUA_Flow;
    private float mopMUA_HX109Flt;
    private float mopMUA_Temp;
    private float mopP201_FltSts;
    private float mopP201_RemSts;
    private float mopP201_RunSts;
    private float mopPabs;
    private float mopPG_FanFault;
    private float mopPG_FanPerm;
    private float mopPG_FanRemote;
    private float mopPG_FanRunH;
    private float mopPG_FanRunL;
    private float mopPG_HL_Tmr;
    private float mopPG_HLActive;
    private float mopPG_HLRunSts;
    private float mopPGDy_FanFlt;
    private float mopPGDy_FanPerm;
    private float mopPGDy_FRunH;
    private float mopPGDy_FRunL;
    private float mopPGDy_HL_Tmr;
    private float mopPGDy_HLRnSts;
    private float mopPGDy_Remote;
    private float mopRelHumidity;
    private float mopRV_Position;
    private float mopSR1_AirTemp;
    private float mopSR1_FaultSts;
    private float mopSR1_Master;
    private float mopSR1_RunSts;
    private float mopSR1_SetptFbk;
    private float mopSR2_AirTemp;
    private float mopSR2_FaultSts;
    private float mopSR2_Master;
    private float mopSR2_RunSts;
    private float mopSR2_SetptFbk;
    private float mopSR3_AirTemp;
    private float mopSR3_FaultSts;
    private float mopSR3_Master;
    private float mopSR3_RunSts;
    private float mopSR3_SetptFbk;
    private float mopSR4_AirTemp;
    private float mopSR4_FaultSts;
    private float mopSR4_Master;
    private float mopSR4_RunSts;
    private float mopSR4_SetptFbk;
    private float mopSR5_AirTemp;
    private float mopSR5_FaultSts;
    private float mopSR5_Master;
    private float mopSR5_RunSts;
    private float mopSR5_SetptFbk;
    private float mopSR6_AirTemp;
    private float mopSR6_FaultSts;
    private float mopSR6_Master;
    private float mopSR6_RunSts;
    private float mopSR6_SetptFbk;
    private float mopSS_ArmExtSts;
    private float mopSS_ArmRetSts;
    private float mopSS_Fan1Fault;
    private float mopSS_Fan1Rem;
    private float mopSS_Fan2Fault;
    private float mopSS_Fan2Rem;
    private float mopSS_FanPerm;
    private float mopSS_Fn1RunSts;
    private float mopSS_Fn2RunSts;
    private float mopSS_HiTemp;
    private float mopSS_InputPyro;
    private float mopSS_Perm;
    private float mopSS_RunSts;
    private float mopSteamPress;
    private float mopSV_Position;
    private float mopTA1_FaultSts;
    private float mopTA2_FaultSts;
    private float mopVH_IgnPerm;
    private float mopVH_IgnSts;
    private int mlDefPriority;
    private float mgvfListParsMissing;
    private float mgvfAT_Previous;
    private float mgvfAT_Sim;
    private float mgvfAT_Sim_002;
    private float mgvfAT_Sim_003;
    private float mgvfAT_Sim_004;
    private float mgvfAT_Sim_005;
    private float mgvfAT_Sim_006;
    private float mgvfCondensTemp;
    private float mgvfDP_Sim;
    private float mgvfDP_Sim_001;
    private float mgvfDP_Sim_002;
    private float mgvfDV_Sim;
    private float mgvfDY_Sim;
    private float mgvfEX_Sim;
    private float mgvfFaultsSim;
    private float mgvfHD_Sim;
    private float mgvfHD_Sim_001;
    private float mgvfHD_Sim_002;
    private float mgvfHD_Sim_003;
    private float mgvfHR_Sim;
    private float mgvfHx_AvgTemp;
    private float mgvfIdle_Sim;
    private float mgvfMF_Previous;
    private float mgvfMF_Sim_001;
    private float mgvfMF_Sim_002;
    private float mgvfMF_SpeedDelayed;
    private float mgvfMF_SpeedDelayType;
    private float mgvfPG_Sim;
    private float mgvfPstatic;
    private float mgvfQ_Sim;
    private float mgvfQ_Sim_001;
    private float mgvfResetSim;
    private float mgvfSF_Sim;
    private float mgvfSR_Sim;
    private float mgvfSS_Sim;
    private float mgvfTimeInterval_Sim;
    private float mgvfTimeInterval_Sim_001;
    private float mgvfWS_Sim;
    private float mgvfWS_Sim_001;
    private Filter1[] mobjFilter1;
    private Ramp1[] mobjRampTowards;
    private Follow[] mobjFollow;
    private Follow2[] mobjFollow2;
    private struct SSLamp
    {
        public float CloudShutterCmd;
        public float CloudShutterFault;
        public float CloudShutterStatus;
        public float Control;
        public float DimmingTime;
        public float FAULT;
        public float LampFault;
        public float Offset;
        public float OffsetSp;
        public float OperatingHours;
        public float OvertempError;
        public float PhaseError;
        public float Power;
        public float Ready;
        public float RestartTime;
        public float Selected;
        public float SelectedStatus;
        public float ShortCircuitError;
        public float Status;
        public float TunnelShutterCmd;
        public float TunnelShutterFault;
        public float TunnelShutterStatus;
    }

    private enum SolarModes
    {
        smOff = 0,
        smManual = 1,
        smAutoPLC = 2,
        smAutoProfile = 3
    }

    private SSLamp[] mudtSSLamp = new SSLamp[28 + 1];
    private void Calc_AT_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        mgvfAT_Previous = mobjRampTowards(0)(mopAirtemp, mgvfAT_Previous, mipAT_Rate / 60, mccfUpdateTime, 0.01);
        if (mgvfAT_Sim == 0 | mbResetSim)
        {
            mopCS_ChSumFlt = 1;
            mopCS_CritFault = 1;
            mopCS_GenFault = 1;
            mopCS_Health = 1;
            mopCS_HtrFltSts = 1;
            mopCS_RunSts = 1;
            mopLTBrn_Temp = -40;
            mopMTBrn_Temp = -25;
            mopHxIn_Temp = 20;
            mopHxOut_Temp = 20;
            mopMixBrn_Temp = mopMTBrn_Temp;
            mopAirtemp = 20;
            mopP201_RemSts = 1;
            mopP201_FltSts = 1;
            mgvfAT_Sim_002 = 0.0001;
            mgvfAT_Sim_003 = 0.1;
            mgvfAT_Sim_004 = 1E-05;
            mgvfAT_Sim_005 = 0.001;
            mgvfAT_Sim_006 = 0.0001;
            mgvfAT_Sim = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        if (mopCS_BrineSts == 1)
        {
            mopCS_LTopnSts = 0;
            mopCS_LTclsSts = mobjFollow2(0)(1, mopCS_LTclsSts, 3);
            mopCS_MTopnSts = mobjFollow2(1)(1, mopCS_MTopnSts, 5);
            mopCS_MTclsSts = 0;
        }
        else if (mopCS_BrineSts == 2)
        {
            mopCS_LTopnSts = mobjRampTowards(1)(1, mopCS_LTopnSts, 5, mccfUpdateTime);
            mopCS_LTclsSts = 0;
            mopCS_MTopnSts = 0;
            mopCS_MTclsSts = mobjRampTowards(2)(1, mopCS_MTclsSts, 3, mccfUpdateTime);
        }

        if (mipCS_HtrSetpt > 0)
        {
            mopCS_HtrRunSts = 1;
        }
        else
        {
            mopCS_HtrRunSts = 0;
        }

        mopCS_BrineSts = mobjFollow2(2)(mipCS_BrineReq, mopCS_BrineSts, 1);
        mopP201_RunSts = mobjFollow2(3)(mipP201_RunReq, mopP201_RunSts, 2);
        mopCV_Position = mipCV_Setpt;
        mopRV_Position = mipRV_Setpt;
        if (mopCS_LTopnSts == 1)
        {
            mopMixBrn_Temp = mobjRampTowards(3)(mopLTBrn_Temp, mopMixBrn_Temp, 20, mccfUpdateTime);
            mopHxIn_Temp = mopHxOut_Temp - mopCV_Position / 100 * mopHxOut_Temp - mopMixBrn_Temp * mgvfAT_Sim_002;
        }
        else if (mopCS_MTopnSts == 1)
        {
            mopMixBrn_Temp = mobjRampTowards(4)(mopMTBrn_Temp, mopMixBrn_Temp, 20, mccfUpdateTime);
            mopHxIn_Temp = mopHxOut_Temp - mopCV_Position / 100 * mopHxOut_Temp - mopMixBrn_Temp * mgvfAT_Sim_003;
        }

        if (mipCS_HtrSetpt != 0)
        {
            mopHxIn_Temp = mopHxIn_Temp + mipCS_HtrSetpt / 100 * 300 * mgvfAT_Sim_004;
        }

        mopHxOut_Temp = mopHxIn_Temp + mgvfAT_Previous - mopHxIn_Temp * mgvfAT_Sim_005;
        mopAirtemp = mgvfAT_Previous + mopHxIn_Temp - mgvfAT_Previous * 0.1 + mopMF_Speed / 950 * 850 * mgvfAT_Sim_006;
        if (mopAirtemp > 55)
        {
            mopAirtemp = 55;
        }

        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        return;
    }

    private void Calc_DV_Sim()
    {
        if (mgvfDV_Sim == 0 | mbResetSim)
        {
            mgvfDV_Sim = 1;
            mopCCE_Perm = 1;
            mopCCE_Remote = 1;
            mopDCA_FaultSts = 1;
            mopDCA_Perm = 1;
            mopDCA_Remote = 1;
            mopDCA_Temp = 21;
            mopDV_Gly_Temp = 70;
            mopDV_Oil_Temp = 50;
            mopDyPit_Temp = 21;
            mopVH_IgnPerm = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        mopCCE_RunSts = mobjFollow2(4)(mipCCE_Start, mopCCE_RunSts, 1);
        mopDCA_RunSts = mobjFollow2(5)(mipDCA_Start, mopDCA_RunSts, 1);
        mopVH_IgnSts = mobjFollow2(6)(mipVH_IgnEnable, mopVH_IgnSts, 1);
        if (mopDCA_RunSts == 1)
        {
            mopDCA_Temp = mobjRampTowards(13)(mipDCA_Target, mopDCA_Temp, 1, mccfUpdateTime);
        }

        if (mipDV_Gly_Start == 1)
        {
            mopDV_Gly_Temp = mobjRampTowards(14)(mipDV_Gly_TTgt, mopDV_Gly_Temp, 1, mccfUpdateTime);
        }

        if (mipDV_Oil_Start == 1)
        {
            mopDV_Oil_Temp = mobjRampTowards(15)(mipDV_Oil_TTgt, mopDV_Oil_Temp, 1, mccfUpdateTime);
        }

        if (mipDyPit_Start == 1)
        {
            mopDyPit_Temp = mobjRampTowards(16)(mipDyPit_Target, mopDyPit_Temp, 1, mccfUpdateTime);
        }

    Done:
        ;
    }

    private void Calc_DY_Sim()
    {
        if (mgvfDY_Sim == 0 | mbResetSim)
        {
            mgvfDY_Sim = 1;
            mopDy_Fault = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

    Done:
        ;
    }

    private void Calc_EX_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mgvfEX_Sim == 0 | mbResetSim)
        {
            mopEX_Fault = 1;
            mopMUA_HX109Flt = 1;
            mopEX_Remote = 1;
            mopEX_Perm = 1;
            mgvfEX_Sim = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        mopEX_RunSts = mobjFollow2(7)(mipEX_Start, mopEX_RunSts, 3);
        if (mopEX_RunSts == 1)
        {
            mopEX_FanSpeed = mobjRampTowards(17)(mipEX_FanSetpt, mopEX_FanSpeed, 400, mccfUpdateTime);
        }
        else
        {
            mopEX_FanSpeed = mobjRampTowards(18)(0, mopEX_FanSpeed, 400, mccfUpdateTime);
        }

        mopBP_Raw = mopEX_FanSpeed * -500 / 1728;
        mopMUA_Temp = mopAirtemp;
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_HD_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        object fHx_AvgTemp = null;
        object fCondensTemp = null;
        if (mgvfHD_Sim == 0 | mbResetSim)
        {
            mopHD_Mirror = 1;
            mopHD_SmplFlow = 1;
            mopSteamPress = 150;
            mopRelHumidity = 5;
            mopMUA_DP = 10;
            mopMUA_Temp = mopAirtemp;
            mopMUA_Flow = 0;
            mgvfHD_Sim_001 = 0.0005;
            mgvfHD_Sim_002 = 0.001;
            mgvfHD_Sim_003 = -0.0015;
            mgvfHD_Sim = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        mopSV_Position = mobjRampTowards(23)(mipSV_Setpt, mopSV_Position, 5, mccfUpdateTime);
        fHx_AvgTemp = mopHxIn_Temp + mopHxOut_Temp / 2;
        if (fHx_AvgTemp < mipDewpoint)
        {
            fCondensTemp = mipDewpoint - fHx_AvgTemp;
        }
        else
        {
            fCondensTemp = 0;
        }

        mopRelHumidity = mopRelHumidity + mopSV_Position * mgvfHD_Sim_001 - fCondensTemp * mgvfHD_Sim_002 + mgvfHD_Sim_003;
        if (mopRelHumidity > 100)
        {
            mopRelHumidity = 100;
        }
        else if (mopRelHumidity < 0)
        {
            mopRelHumidity = 0;
        }

        mopMUA_DP = mobjRampTowards(24)(mipDP_Target, mopMUA_DP, 0.5, mccfUpdateTime, 0.1);
        mopMUA_Temp = mopAirtemp;
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_HR_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mgvfHR_Sim == 0 | mbResetSim)
        {
            mgvfHR_Sim = 1;
            mopHR_FaultSts = 1;
            mopHR_Perm = 1;
            mopHR_EstopSts = 1;
            mopHR1_FaultSts = 1;
            mopHR2_FaultSts = 1;
            mopHR3_FaultSts = 1;
            mopHR4_FaultSts = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        if (mipHR_RunSts == 1)
        {
            mopHR_Brn_Temp = mobjRampTowards(31)(mipHR_BrnSetpt, mopHR_Brn_Temp, 0.5, mccfUpdateTime);
        }

        for (i = 1; i <= 4; i++)
        {
            switch (i)
            {
                case 1:
                    fHR_RunSts = mipHR1_RunSts;
                    fHR_HxInTmp = mopHR1_HxInTmp;
                    fHR_HxOutTmp = mopHR1_HxOutTmp;
                    fHR_CVSetpt = mipHR1_CVSetpt;
                    fHR_HtrSetpt = mipHR1_HtrSetpt;
                    fHR_PnlTemp = mopHR1_PnlTemp;
                    break;
                case 2:
                    fHR_RunSts = mipHR2_RunSts;
                    fHR_HxInTmp = mopHR2_HxInTmp;
                    fHR_HxOutTmp = mopHR2_HxOutTmp;
                    fHR_CVSetpt = mipHR2_CVSetpt;
                    fHR_HtrSetpt = mipHR2_HtrSetpt;
                    fHR_PnlTemp = mopHR2_PnlTemp;
                    break;
                case 3:
                    fHR_RunSts = mipHR3_RunSts;
                    fHR_HxInTmp = mopHR3_HxInTmp;
                    fHR_HxOutTmp = mopHR3_HxOutTmp;
                    fHR_CVSetpt = mipHR3_CVSetpt;
                    fHR_HtrSetpt = mipHR3_HtrSetpt;
                    fHR_PnlTemp = mopHR3_PnlTemp;
                    break;
                case 4:
                    fHR_RunSts = mipHR4_RunSts;
                    fHR_HxInTmp = mopHR4_HxInTmp;
                    fHR_HxOutTmp = mopHR4_HxOutTmp;
                    fHR_CVSetpt = mipHR4_CVSetpt;
                    fHR_HtrSetpt = mipHR4_HtrSetpt;
                    fHR_PnlTemp = mopHR4_PnlTemp;
                    break;
            }

            if (fHR_RunSts == 1)
            {
                fHR_HxInTmp = fHR_HxOutTmp - fHR_CVSetpt / 100 * fHR_HxOutTmp - mopHR_Brn_Temp;
                if (fHR_HtrSetpt != 0)
                {
                    fHR_HxInTmp = fHR_HxInTmp + fHR_HtrSetpt / 100 * 21;
                }

                fHR_HxOutTmp = fHR_HxInTmp + fHR_PnlTemp - fHR_HxInTmp;
                fHR_PnlTemp = fHR_PnlTemp + fHR_HxInTmp - fHR_PnlTemp * 0.1;
                if (fHR_PnlTemp > 80)
                {
                    fHR_PnlTemp = 80;
                }
                else if (fHR_PnlTemp < 10)
                {
                    fHR_PnlTemp = 10;
                }
            }
            else
            {
                fHR_HxInTmp = mobjRampTowards(32)(mopAirtemp, fHR_HxInTmp, 0.5, mccfUpdateTime);
                fHR_HxOutTmp = fHR_HxInTmp;
                fHR_PnlTemp = fHR_HxInTmp;
            }

            switch (i)
            {
                case 1:
                    mopHR1_HxInTmp = fHR_HxInTmp;
                    mopHR1_HxOutTmp = fHR_HxOutTmp;
                    mopHR1_PnlTemp = fHR_PnlTemp;
                    break;
                case 2:
                    mopHR2_HxInTmp = fHR_HxInTmp;
                    mopHR2_HxOutTmp = fHR_HxOutTmp;
                    mopHR2_PnlTemp = fHR_PnlTemp;
                    break;
                case 3:
                    mopHR3_HxInTmp = fHR_HxInTmp;
                    mopHR3_HxOutTmp = fHR_HxOutTmp;
                    mopHR3_PnlTemp = fHR_PnlTemp;
                    break;
                case 4:
                    mopHR4_HxInTmp = fHR_HxInTmp;
                    mopHR4_HxOutTmp = fHR_HxOutTmp;
                    mopHR4_PnlTemp = fHR_PnlTemp;
                    break;
            }
        }

        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_IC_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mgvfIdle_Sim == 0 | mbResetSim)
        {
            mgvfIdle_Sim = 1;
            mopIDL_Fault = 1;
            mopIDL_Perm = 1;
            mopIDL_Remote = 1;
            mopIDL_ModeOn = 0;
            mopIDL_ModeOff = 1;
            mopIDL_FlOpdSts = 0;
            mopIDL_FlCldSts = 1;
            mopIDL_ShCldSts = 0;
            mopIDL_ShOpdSts = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        if (mipIDL_Start == 1)
        {
            mopIDL_ModeOn = 1;
            mopIDL_ModeOff = 0;
            if (mopIDL_FlOpdSts == 0)
            {
                mopIDL_FlOpdSts = mobjFollow2(8)(1, mopIDL_FlOpdSts, 5);
                mopIDL_FlCldSts = 0;
            }
            else if (mopIDL_ShCldSts == 0)
            {
                mopIDL_ShCldSts = mobjFollow2(9)(1, mopIDL_ShCldSts, 5);
                mopIDL_ShOpdSts = 0;
            }
        }
        else
        {
            if (mopIDL_ShOpdSts == 0)
            {
                mopIDL_ShOpdSts = mobjFollow2(10)(1, mopIDL_ShOpdSts, 5);
                mopIDL_ShCldSts = 0;
            }
            else if (mopIDL_FlCldSts == 0)
            {
                mopIDL_FlCldSts = mobjFollow2(11)(1, mopIDL_FlCldSts, 5);
                mopIDL_FlOpdSts = 0;
            }
            else
            {
                mopIDL_ModeOn = 0;
                mopIDL_ModeOff = 1;
            }
        }

        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_Faults_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mgvfFaultsSim == 0 | mbResetSim)
        {
            mgvfFaultsSim = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_MF_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_PG_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mgvfPG_Sim == 0 | mbResetSim)
        {
            mopPG_FanFault = 1;
            mopPG_FanPerm = 1;
            mopPG_FanRemote = 1;
            mopPG_HLActive = 1;
            mopPGDy_FanFlt = 1;
            mopPGDy_FanPerm = 1;
            mopPGDy_Remote = 1;
            mgvfPG_Sim = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        if (mipPG_HL_Start == 1)
        {
            if (mipPG_HL_DelRst == 1)
            {
                mopPG_HL_Tmr = 30;
            }
            else
            {
                mopPG_HL_Tmr = mobjRampTowards(45)(0, mopPG_HL_Tmr, 1, mccfUpdateTime);
            }

            if (mopPG_HL_Tmr > 0)
            {
            }
            else
            {
                mopPG_FanRunH = mobjFollow2(12)(mipPG_HL_Start, mopPG_FanRunH, 1);
            }
        }
        else
        {
            mopPG_HL_Tmr = 30;
            mopPG_FanRunH = mobjFollow2(13)(mipPG_HL_Start, mopPG_FanRunH, 1);
        }

        mopPG_HLRunSts = mopPG_FanRunH;
        mopPG_FanRunL = mobjFollow2(14)(mipPG_LL_Start, mopPG_FanRunL, 1);
        if (mipPGDy_HL_Strt == 1)
        {
            if (mipPGDy_HL_DRst == 1)
            {
                mopPGDy_HL_Tmr = 30;
            }
            else
            {
                mopPGDy_HL_Tmr = mobjRampTowards(46)(0, mopPGDy_HL_Tmr, 1, mccfUpdateTime);
            }

            if (mopPGDy_HL_Tmr > 0)
            {
            }
            else
            {
                mopPGDy_FRunH = mobjFollow2(15)(mipPGDy_HL_Strt, mopPGDy_FRunH, 1);
            }
        }
        else
        {
            mopPGDy_HL_Tmr = 30;
            mopPGDy_FRunH = mobjFollow2(16)(mipPGDy_HL_Strt, mopPGDy_FRunH, 1);
        }

        mopPGDy_HLRnSts = mopPGDy_FRunH;
        mopPGDy_FRunL = mobjFollow2(17)(mipPGDy_LL_Strt, mopPGDy_FRunL, 1);
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_SF_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mgvfSF_Sim == 0 | mbResetSim)
        {
            mgvfSF_Sim = 1;
            mopCtrl_Power = 1;
            mopDoor_UpLg_AS = 1;
            mopDoor_Dy = 1;
            mopKey_Veh_Door = 1;
            mopKey_Dy_Door = 1;
            mopFire_Sum = 1;
            mopFire_DYSup = 1;
            mopLoLo_Plen_P = 1;
            mopLo_Plen_P = 1;
            mopHX_HiDeltaP = 1;
            mopEPB_TS_NW = 1;
            mopEPB_TS_SW = 1;
            mopEPB_UpLeg_AS = 1;
            mopEStop_SumSts = 1;
            mopEPB_Ctrl_Rm = 1;
            mopFaultBus = 1;
            mopDy_Estop_Sts = 1;
            mopGas_Sys_Sts = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_SR_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mgvfSR_Sim == 0 | mbResetSim)
        {
            mgvfSR_Sim = 1;
            mopAL1_FaultSts = 1;
            mopAL2_FaultSts = 1;
            mopSR1_FaultSts = 1;
            mopSR2_FaultSts = 1;
            mopSR3_FaultSts = 1;
            mopSR4_FaultSts = 1;
            mopSR5_FaultSts = 1;
            mopSR6_FaultSts = 1;
            mopTA1_FaultSts = 1;
            mopTA2_FaultSts = 1;
            mopSR1_Master = 2;
            mopSR2_Master = mopSR1_Master;
            mopSR3_Master = mopSR1_Master;
            mopSR4_Master = mopSR1_Master;
            mopSR5_Master = mopSR1_Master;
            mopSR6_Master = mopSR1_Master;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        mopSR1_SetptFbk = mipSR1_Setpt;
        mopSR2_SetptFbk = mipSR2_Setpt;
        mopSR3_SetptFbk = mipSR3_Setpt;
        mopSR4_SetptFbk = mipSR4_Setpt;
        mopSR5_SetptFbk = mipSR5_Setpt;
        mopSR6_SetptFbk = mipSR6_Setpt;
        mopSR1_AirTemp = mobjRampTowards(47)(mipSR1_Setpt, mopSR1_AirTemp, 2, mccfUpdateTime);
        mopSR2_AirTemp = mobjRampTowards(48)(mipSR2_Setpt, mopSR2_AirTemp, 2, mccfUpdateTime);
        mopSR3_AirTemp = mobjRampTowards(49)(mipSR3_Setpt, mopSR3_AirTemp, 2, mccfUpdateTime);
        mopSR4_AirTemp = mobjRampTowards(50)(mipSR4_Setpt, mopSR4_AirTemp, 2, mccfUpdateTime);
        mopSR5_AirTemp = mobjRampTowards(51)(mipSR5_Setpt, mopSR5_AirTemp, 2, mccfUpdateTime);
        mopSR6_AirTemp = mobjRampTowards(52)(mipSR6_Setpt, mopSR6_AirTemp, 2, mccfUpdateTime);
        mopSR1_RunSts = mobjFollow2(18)(mipSR1_RunReq, mopSR1_RunSts, 1);
        mopSR2_RunSts = mobjFollow2(19)(mipSR2_RunReq, mopSR2_RunSts, 1);
        mopSR3_RunSts = mobjFollow2(20)(mipSR3_RunReq, mopSR3_RunSts, 1);
        mopSR4_RunSts = mobjFollow2(21)(mipSR4_RunReq, mopSR4_RunSts, 1);
        mopSR5_RunSts = mobjFollow2(22)(mipSR5_RunReq, mopSR5_RunSts, 1);
        mopSR6_RunSts = mobjFollow2(23)(mipSR6_RunReq, mopSR6_RunSts, 1);
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_SS_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mgvfSS_Sim == 0 | mbResetSim)
        {
            mgvfSS_Sim = 1;
            mopSS_ArmRetSts = 1;
            mopSS_Fan1Fault = 1;
            mopSS_Fan1Rem = 1;
            mopSS_Fan2Fault = 1;
            mopSS_Fan2Rem = 1;
            mopSS_FanPerm = 1;
            mopSS_HiTemp = 1;
            mopSS_Perm = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        mopSS_ArmRetSts = mobjFollow2(24)(-mipSS_ArmExtReq + 1, mopSS_ArmRetSts, 5);
        mopSS_ArmExtSts = mobjFollow2(25)(mipSS_ArmExtReq, mopSS_ArmExtSts, 5);
        if (mipSS_ArmExtReq == 0)
        {
            mopSS_ArmExtSts = 0;
        }
        else
        {
            mopSS_ArmRetSts = 0;
        }

        mopSS_InputPyro = mipSS_InvCoef0 - 28.4 + mipSS_InvCoef1 * mipSS_Power + mipSS_InvCoef2 * Math.Pow(mipSS_Power, 2) + mipSS_InvCoef3 * Math.Pow(mipSS_Power, 3) + mipSS_InvCoef4 * Math.Pow(mipSS_Power, 4);
        mopSS_Fn1RunSts = mobjRampTowards(53)(mipSS_FanStart, mopSS_Fn1RunSts, 1, mccfUpdateTime);
        mopSS_Fn2RunSts = mobjRampTowards(54)(mipSS_FanStart, mopSS_Fn2RunSts, 1, mccfUpdateTime);
        mopSS_RunSts = mobjRampTowards(55)(mipSS_RunReq, mopSS_RunSts, 1, mccfUpdateTime);
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_TimeInterval_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        CurrentSeconds = Timer;
        mgvfTimeInterval_Sim = CurrentSeconds - mgvfTimeInterval_Sim_001;
        mgvfTimeInterval_Sim = Ctl_Limit(mgvfTimeInterval_Sim, 0.05, 1);
        mgvfTimeInterval_Sim_001 = CurrentSeconds;
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void Calc_WS_Sim()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mgvfWS_Sim == 0 | mbResetSim)
        {
            mopMF_Remote = 1;
            mopMF_Fault = 1;
            mopMF_Perm = 1;
            mopMF_SupplySts = 1;
            mopMF_Warning = 1;
            mopMF_SpdLim1 = 0;
            mopMF_SpdLim2 = 0;
            mgvfMF_Sim_001 = 0;
            mgvfMF_Sim_002 = 0.2;
            mgvfQ_Sim_001 = 1;
            mgvfDP_Sim_001 = 1;
            mgvfDP_Sim_002 = 0;
            mgvfWS_Sim_001 = 0.2;
            mopPabs = 101.325;
            mgvfWS_Sim = 1;
            if (mbResetSim)
            {
                goto Done;
            }
        }

        switch (mgvfMF_SpeedDelayType)
        {
            case 1:
                mfMFSpeedSPBuffer(mlMFSpeedSPBufferIndex) = mipMF_Setpt;
                mlMFSpeedSPBufferIndex = mlMFSpeedSPBufferIndex + 1;
                if (mlMFSpeedSPBufferIndex > mlMFSpeedSPBufferSize)
                {
                    mlMFSpeedSPBufferIndex = 0;
                }

                mopMF_Speed = mfMFSpeedSPBuffer(mlMFSpeedSPBufferIndex);
                break;
            default:
                mopMF_Speed = mobjFilter1(8)(mipMF_Setpt, mgvfMF_Sim_001);
                break;
        }

        mgvfPstatic = mopPabs * 1000;
        mgvfMF_SpeedDelayed = mobjFilter1(9)(mopMF_Speed, mgvfMF_Sim_002);
        fWindspeed = mgvfMF_SpeedDelayed / 1000 * 230 / 3.6;
        mgvfQ_Sim = 0.5 * mipAirDensity * Math.Pow(fWindspeed, 2);
        mgvfQ_Sim = mgvfQ_Sim * mgvfQ_Sim_001;
        mgvfDP_Sim = mgvfPstatic * Math.Pow(1 + mgvfQ_Sim / 3.5 * mgvfPstatic, 3.5) - 1;
        mgvfDP_Sim = mgvfDP_Sim * mgvfDP_Sim_001;
        mgvfDP_Sim = mobjFilter1(11)(mgvfDP_Sim, mgvfDP_Sim_002);
        mopDeltaP_Lo = Ctl_Limit(mgvfDP_Sim, 0, mccfDPxOverHi);
        mopDeltaP_Hi = Ctl_Limit(mgvfDP_Sim, 0, 2500);
        if (mipMF_Start == 0)
        {
            mopMF_RunSts = mobjFollow2(26)(mipMF_Start, mopMF_RunSts, 0.5);
        }
        else
        {
            mopMF_RunSts = mobjFollow2(27)(mipMF_Start, mopMF_RunSts, 5);
        }

        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        return;
    }

    private void CalcDynamicPressurePa(ref float DeltaPPa, ref float AbsPkPa, ref float Kp, ref float Kq, ref float BlockageFactor, ref float CalculatedPStatickPa, ref float CalculatedQPa)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        PStatic = AbsPkPa * 1000 + Kp * DeltaPPa;
        Pt = PStatic + Kq * DeltaPPa;
        CalculatedQPa = 1 / K1 * PStatic * Math.Pow(Pt / PStatic, K1) - 1 * BlockageFactor;
        if (CalculatedQPa < 0)
        {
            CalculatedQPa = 0;
        }

        CalculatedPStatickPa = PStatic / 1000;
        goto Done;
        Error();
        CalculatedQPa = 0;
        CalculatedPStatickPa = AbsPkPa;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, TypeName(this), PROC_NAME, sRaiseErrorText);
        }

        return;
    }

    public void CheckInterlock(ref object Opcode, object CheckValue = null)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        // On Error Resume Next — nested handler; see enclosing try/catch
        sIndex = Opcode;
        objInterlock = mobjInterlocks(sIndex);
        if (Err)
        {
            goto Done;
        }

        mvInterlockCheckValue = CheckValue;
        // On Error GoTo Error — nested handler not restructured
        objInterlock.Check();
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        lRaiseErrorNumber = ex.HResult;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    private void Class_Initialize()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (App.PrevInstance)
        {
            if (App.StartMode != vbSModeStandalone)
            {
                EH.AppError = StdTalentStrings.stsEHInvalidActiveXLaunch;
                EH.Item(1) = PROCESS_NAME;
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }

            End();
        }// On Error GoTo Error — nested handler not restructured
        App.StartLogging("", vbLogAuto);
        App.OleServerBusyRaiseError = true;
        App.OleServerBusyTimeout = 1000;
        Process.Title = PROCESS_NAME;
        Process.EHCaption = PROCESS_NAME;
        Process.EHMode = ehLogOnly;
        Process.EHMsgsToLog = ehCriticalExclamation;
        Process.LogMode = logToTalentAndFile;
        // On Error GoTo Error_CreatingObjects — nested handler not restructured
        sClassName = "LoopTimer";
        tmrMain = new LoopTimer();
        if (App.StartMode != vbSModeStandalone)
        {
            tmrRunOnce = new ccrpTimers6.ccrpTimer();
        }// On Error GoTo Error — nested handler not restructured
        if (App.StartMode != vbSModeStandalone)
        {
            tmrRunOnce.Interval = 500;
            tmrRunOnce.EventType = TimerOneShot;
            tmrRunOnce.Enabled = true;
        }

        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto Done;
    Error_CreatingObjects:
        ;
        sRaiseErrorText = string.Concat(LoadResString(ridCantCreateObject, sClassName), ex.Message);
        bRaiseError = true;
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        EH.Item(1) = sClassName;
        EH_MsgBox();
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    private void Class_Terminate()
    {
        int i = 0;
        try
        {
            tmrMain.Enabled = false;
            tmrMain.Notify = null;
            tmrMain = null;
            for (i = 1; i <= mlNumFilter1; i++)
            {
                mobjFilter1(i) = null;
            }

            for (i = 1; i <= mlNumRampTowards; i++)
            {
                mobjRampTowards(i) = null;
            }

            for (i = 1; i <= mlNumFollow; i++)
            {
                mobjFollow(i) = null;
            }

            for (i = 1; i <= mlNumFollow2; i++)
            {
                mobjFollow2(i) = null;
            }

            tmrElapsed = null;
            mobjCommands = null;
            mobjConstants = null;
            mobjGlobalVars = null;
            mobjEvents = null;
            mobjStats = null;
            mobjStatuses = null;
            mcolInvalidOpcodes = null;
            mobjSPGConstants = null;
            mobjSPGDataItems = null;
            mobjParManager.Release();
            mobjParManager = null;
            mbInitialized = false;
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    public CommandListLib.Commands Commands
    {
        get
        {
            CommandListLib.Commands __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = mobjCommands;
            goto Done;
            Error();
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }

            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
            }

            return __result;
        }
    }

    public ControlLib.ControlConstants ControlConstants
    {
        get
        {
            ControlLib.ControlConstants __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = mobjConstants;
            goto Done;
            Error();
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
            }

            return __result;
        }
    }

    internal void ControlLogic()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        ExecuteCommands();
        if (mbInitialized)
        {
            mobjParManager.ReadInputs();
            if (mgvfAT_Sim == 0 | mgvfResetSim == 0)
            {
                mbResetSim = true;
            }

            Calc_TimeInterval_Sim();
            Calc_AT_Sim();
            Calc_DV_Sim();
            Calc_DY_Sim();
            Calc_EX_Sim();
            Calc_Faults_Sim();
            Calc_HD_Sim();
            Calc_HR_Sim();
            Calc_IC_Sim();
            Calc_MF_Sim();
            Calc_PG_Sim();
            Calc_SF_Sim();
            Calc_SR_Sim();
            Calc_SS_Sim();
            Calc_WS_Sim();
            if (mbResetSim)
            {
                mbResetSim = false;
                mgvfResetSim = 1;
            }

            mobjParManager.WriteOutputs();
        }

        if (!mobjStatuses == null)
        {
            mobjStatuses.FlushMessages();
        }

        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        Suspend();
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    public Events Events
    {
        get
        {
            Events __result = 0;
            __result = mobjEvents;
            return __result;
        }
    }

    private void ExecuteCommands()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        if (ReEntered)
        {
            return;
        }

        ReEntered = true;
        if (mobjCommands.Count == 0)
        {
            goto Done;
        }

        objCommand = mobjCommands(1);
        switch (objCommand.Opcode)
        {
            case ocInitialize:
                this.Initialize(,);
                objCommand();
                objCommand.Remove();
                mobjStatuses(STS_SIM_INITIALIZED).SetValue(mbInitialized);
                break;
            case ocSuspend:
                objCommand.Messages.Flush();
                mbInitialized = false;
                objCommand.Remove();
                mobjStatuses(STS_SIM_INITIALIZED).SetValue(mbInitialized);
                SetProcessPriority(priNormal);
                break;
            case ocShutdown:
                objCommand.Messages.Flush();
                objCommand.Remove();
                this.Shutdown();
                break;
            default:
                CheckInterlock(ocInitialize);
                switch (objCommand.Opcode)
                {
                    default:
                        objCommand.Abort(string.Concat("Unrecognized opcode: ", objCommand.Opcode));
                        break;
                }

                break;
        }

        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto ErrorExit;
    ErrorCreatingSocket:
        ;
        sRaiseErrorText = ex.Message;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        EH_MsgBox();
        goto ErrorExit;
    ErrorExit:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, objCommand.Text, sRaiseErrorText);
        objCommand.Abort(sRaiseErrorText, cmdAborted);
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        ReEntered = false;
        return;
    }

    public object GetInvalidOpcodes(ref object Opcode)
    {
        object __result = null;
        string sIndex = "";
        try
        {
            sIndex = Opcode;
            __result = mcolInvalidOpcodes(sIndex);
            if (Err)
            {
                __result = 0;
            }

            return __result;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }

        return __result;
    }

    public ControlLib.GlobalVars GlobalVars
    {
        get
        {
            ControlLib.GlobalVars __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = mobjGlobalVars;
            goto Done;
            Error();
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
            }

            return __result;
        }
    }

    private void ICcrpTimerNotify_Timer(int Milliseconds)
    {
        int lElapsed = 0;
        try
        {
            ControlLogic();
            if (mbKeepStats & mbInitialized)
            {
                lElapsed = tmrElapsed.Elapsed;
                tmrElapsed.Reset();
                mobjStats.AddValue((double)lElapsed / 1000);
            }

            DoEvents();
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    internal void Initialize(InitTypes InitType = initAll, Command Command = 0)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        switch (InitType)
        {
            case initConstants:
            case initAll:
                bStatus = true;
                sPath = GetDefaultCCFolder(App.Path);
                sFileName = string.Concat(string.Concat(sPath, PROCESS_NAME), ".csv");
                sDefaultsFileName = string.Concat(string.Concat(string.Concat(App.Path, "\\"), App.EXEName), "Defaults.csv");
                mobjConstants.OpenFileWithDefaults(sFileName, sDefaultsFileName);
                mobjConstants.AddString("Priority", mccsPriority, bStatus, sText);
                mobjConstants.AddSingle("UpdateInterval", mccfUpdateTime, bStatus, sText);
                mobjConstants.AddSingle("FanDelay", mccfFanDelay, bStatus, sText);
                mobjConstants.AddSingle("TunnelVolume", mccfTunnelVolume, bStatus, sText);
                if (bStatus == false)
                {
                    bRaiseError = true;
                    sRaiseErrorText = string.Concat(string.Concat(string.Concat("Can't find the following control constants in ", sFileName), ": "), sText);
                    lRaiseErrorNumber = tceCommandFailed;
                    if (Command == null)
                    {
                    }
                    else
                    {
                        Command.SuppressLogging = true;
                    }

                    objEH = new EH();
                    objEH.AppError = LoadResString(StdTalentStrings.stsEHCantLoadConstants);
                    objEH.Item(1) = sText;
                    objEH.FileName = sFileName;
                    objEH.Module = TypeName(this);
                    objEH.Procedure = PROC_NAME;
                    LogMessage2(objEH);
                }

                mobjSPGConstants.CreatePropertyByPosition(ccUpdateTime, bExists).Value = mccfUpdateTime;
                objConstants = new ControlLib.ControlConstants();
                objConstants.OpenFile(string.Concat(sPath, "WindSpeed.csv"));
                objConstants.AddSingle("KP1_C0", mccfKp1Coeff(0), bStatus, sText);
                objConstants.AddSingle("KP1_C1", mccfKp1Coeff(1), bStatus, sText);
                objConstants.AddSingle("KP1_C2", mccfKp1Coeff(2), bStatus, sText);
                objConstants.AddSingle("KP2_C0", mccfKp2Coeff(0), bStatus, sText);
                objConstants.AddSingle("KP2_C1", mccfKp2Coeff(1), bStatus, sText);
                objConstants.AddSingle("KP2_C2", mccfKp2Coeff(2), bStatus, sText);
                objConstants.AddSingle("KP_DPTRANS", mccfKpDpTrans, bStatus, sText);
                objConstants.AddSingle("KQ1_C0", mccfKq1Coeff(0), bStatus, sText);
                objConstants.AddSingle("KQ1_C1", mccfKq1Coeff(1), bStatus, sText);
                objConstants.AddSingle("KQ1_C2", mccfKq1Coeff(2), bStatus, sText);
                objConstants.AddSingle("KQ2_C0", mccfKq2Coeff(0), bStatus, sText);
                objConstants.AddSingle("KQ2_C1", mccfKq2Coeff(1), bStatus, sText);
                objConstants.AddSingle("KQ2_C2", mccfKq2Coeff(2), bStatus, sText);
                objConstants.AddSingle("KQ_DPTRANS", mccfKqDpTrans, bStatus, sText);
                objConstants.AddSingle("DPXHI", mccfDPxOverHi, bStatus, sText);
                objConstants.AddSingle("DPXLO", mccfDPxOverLo, bStatus, sText);
                objConstants = null;
                break;
        }

        mlDefPriority = GetControllerPriority(mccsPriority);
        mbResetSim = true;
        if (mgvfListParsMissing == 0)
        {
            // On Error Resume Next — nested handler; see enclosing try/catch
            mobjParManager.RefreshObjects();
            // On Error GoTo Error — nested handler not restructured
            mbInitialized = true;
        }
        else
        {
            switch (InitType)
            {
                case initAll:
                case initPars:
                    // On Error Resume Next — nested handler; see enclosing try/catch
                    ;
                    mobjParManager.RefreshObjects();
                    if (Err)
                    {
                        lRaiseErrorNumber = CCommandError(ex.HResult);
                        sRaiseErrorText = ex.Message;
                        objEH = new EH();
                        objEH.VBErr = Err;
                        objEH.AppError = LoadResString(StdTalentStrings.stsEHInitializeFailedParsMissing);
                        objEH.Item(1) = Command.Text;
                        objEH.Module = TypeName(this);
                        objEH.Procedure = PROC_NAME;
                        if (Command == null)
                        {
                        }
                        else
                        {
                            objEH.Item(1) = Command.Text;
                        }

                        LogMessage2(objEH);
                        if (Command == null)
                        {
                        }
                        else
                        {
                            Command.SuppressLogging = true;
                        }

                        goto Done;
                    }

                    mbInitialized = mobjParManager.Ready;
                    break;
            }
        }

        for (i = 0; i <= mlNumFilter1; i++)
        {
            mobjFilter1(i) = new Filter1();
        }

        for (i = 0; i <= mlNumRampTowards; i++)
        {
            mobjRampTowards(i) = new Ramp1();
        }

        for (i = 0; i <= mlNumFollow; i++)
        {
            mobjFollow(i) = new Follow();
        }

        for (i = 0; i <= mlNumFollow2; i++)
        {
            mobjFollow2(i) = new Follow2();
        }

        mlMFSpeedSPBufferSize = (int)mccfFanDelay / mccfUpdateTime;
        mfMFSpeedSPBuffer = new object[mlMFSpeedSPBufferSize + 1];
        for (i = 0; i <= mlMFSpeedSPBufferSize; i++)
        {
            mfMFSpeedSPBuffer(i) = mipMF_Setpt;
            mlMFSpeedSPBufferIndex = 0;
        }

        goto Done;
        Error();
        mbInitialized = false;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            if (Command == null)
            {
            }
            else
            {
                EH.Item(1) = Command.Text;
            }

            EH_MsgBox();
            if (Command == null)
            {
            }
            else
            {
                Command.SuppressLogging = true;
            }
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        mobjStatuses(STS_SIM_INITIALIZED).SetValue(mbInitialized);
        mobjConstants.CloseFile();
        objEH = null;
        if (mbInitialized)
        {
            SetProcessPriority(mlDefPriority);
        }
        else
        {
            SetProcessPriority(priNormal);
        }

        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    public bool Initialized
    {
        get
        {
            bool __result = false;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = mbInitialized;
            goto Done;
            Error();
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }

            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText);
            }

            return __result;
        }
    }

    public Interlocks Interlocks
    {
        get
        {
            Interlocks __result = 0;
            __result = mobjInterlocks;
            return __result;
        }
    }

    private void LogMessageEH(ref string MessageText, ref string SendingRoutine, object Items = null, MessageTypes MessageType = msgInformation)
    {
        EH objEH = 0;
        string[] sItems;
        int i = 0;
        try
        {
            objEH = new EH();
            objEH.AppError = MessageText;
            if (!IsMissing(Items))
            {
                for (i = 1; i <= 10; i++)
                {
                    // Err.Clear()
                    objEH.Item(i) = Items(i);
                    if (Err)
                    {
                        break;
                    }
                }
            }

            objEH.Module = TypeName(this);
            objEH.Procedure = SendingRoutine;
            objEH.MessageType = MessageType;
            LogMessage2(objEH);
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    public void Main()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        App.StartLogging("", vbLogAuto);
        App.OleServerBusyRaiseError = true;
        App.OleServerBusyTimeout = 1000;
        Process.Title = PROCESS_NAME;
        Process.EHCaption = PROCESS_NAME;
        Process.EHMode = ehLogOnly;
        Process.EHMsgsToLog = ehCriticalExclamation;
        Process.LogMode = logToTalentAndFile;
        goto Done;
        Error();
        if (Err != 364)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto ErrorExit;
    ErrorExit:
        ;
        EH.AppError = ERR_CANT_START;
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        EH_MsgBox();
        // On Error Resume Next — nested handler; see enclosing try/catch
        goto Done;
    Done:
        ;
        return;
    }

    public object ParManager
    {
        get
        {
            object __result = null;
            __result = mobjParManager;
            return __result;
        }
    }

    public bool Ready
    {
        get
        {
            bool __result = false;
            try
            {
                __result = mbReady;
                return __result;
            }
            catch
            {
                // On Error Resume Next: exceptions suppressed
            }

            return __result;
        }
    }

    public void Shutdown()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        SetProcessPriority(priNormal);
        mbReady = false;
        mobjCommands.Clear("Controller shutting down");
        DoEvents();
        mbInitialized = false;
        mobjParManager.ReleaseOutputs();
        mobjEvents.RaiseQuit();
        DoEvents();
        Unload(mobjForm);
        mobjForm = null;
        objSOM = new SOM();
        objSOM.RemoveObject(SOM_CONTROLLER);
        objSOM = null;
        Class_Terminate();
        Pause(2);
        if (IsWithinIDE())
        {
            End();
        }
        else
        {
            TerminateProcess(GetCurrentProcess(), 0);
        }

        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        EH_MsgBox();
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    private void Start()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        // On Error GoTo Error_CreatingObjects — nested handler not restructured
        sClassName = "CommandListLib.Commands";
        mobjCommands = new CommandListLib.Commands();
        sClassName = "ControlLib.ControlConstants";
        mobjConstants = new ControlLib.ControlConstants();
        sClassName = "Events";
        mobjEvents = new Events();
        sClassName = "Interlocks";
        mobjInterlocks = new Interlocks();
        sClassName = "Invalid Opcodes Collection";
        mcolInvalidOpcodes = new Collection();
        sClassName = "frmControl";
        mobjForm = new frmControl();
        sClassName = "SharedPropertyGroupManager";
        objSPGM = new SharedPropertyGroupManager();
        mobjSPGConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists);
        sClassName = "Function Objects";
        mlNumFilter1 = 11;
        mobjFilter1 = new Filter1[mlNumFilter1 + 1];
        for (i = 0; i <= mlNumFilter1; i++)
        {
            mobjFilter1(i) = new Filter1();
        }

        mlNumRampTowards = 68;
        mobjRampTowards = new Ramp1[mlNumRampTowards + 1];
        for (i = 0; i <= mlNumRampTowards; i++)
        {
            mobjRampTowards(i) = new Ramp1();
        }

        mlNumFollow = 30;
        mobjFollow = new Follow[mlNumFollow + 1];
        for (i = 0; i <= mlNumFollow; i++)
        {
            mobjFollow(i) = new Follow();
        }

        mlNumFollow2 = 27;
        mobjFollow2 = new Follow2[mlNumFollow2 + 1];
        for (i = 0; i <= mlNumFollow2; i++)
        {
            mobjFollow2(i) = new Follow2();
        }// On Error GoTo Error — nested handler not restructured
        mobjStats = new AdvUtilLib.Stats();
        tmrElapsed = new ccrpStopWatch();
        DoEvents();
        mobjInterlocks.Add(ocInitialize.ToString()).Expressions.AddBoolean("Initialized", mbInitialized, true, ilckVariableEqualValueMeansOK, LoadResString(stsControlSystemNotInitialized));
        // On Error Resume Next — nested handler; see enclosing try/catch
        mobjGlobalVars = new GlobalVars();
        mobjGlobalVars.AddBoolean("Initialized", mbInitialized);
        mobjGlobalVars.AddSingle("mgvfListParsMissing 1=True", mgvfListParsMissing);
        mobjGlobalVars.AddSingle("AT_Previous", mgvfAT_Previous);
        mobjGlobalVars.AddSingle("AT_Sim", mgvfAT_Sim);
        mobjGlobalVars.AddSingle("AT_Sim_002", mgvfAT_Sim_002);
        mobjGlobalVars.AddSingle("AT_Sim_003", mgvfAT_Sim_003);
        mobjGlobalVars.AddSingle("AT_Sim_004", mgvfAT_Sim_004);
        mobjGlobalVars.AddSingle("AT_Sim_005", mgvfAT_Sim_005);
        mobjGlobalVars.AddSingle("AT_Sim_006", mgvfAT_Sim_006);
        mobjGlobalVars.AddSingle("CondensTemp", mgvfCondensTemp);
        mobjGlobalVars.AddSingle("DP_Sim", mgvfDP_Sim);
        mobjGlobalVars.AddSingle("DP_Sim_001", mgvfDP_Sim_001);
        mobjGlobalVars.AddSingle("DP_Sim_002", mgvfDP_Sim_002);
        mobjGlobalVars.AddSingle("DV_Sim", mgvfDV_Sim);
        mobjGlobalVars.AddSingle("DY_Sim", mgvfDY_Sim);
        mobjGlobalVars.AddSingle("EX_Sim", mgvfEX_Sim);
        mobjGlobalVars.AddSingle("FaultsSim", mgvfFaultsSim);
        mobjGlobalVars.AddSingle("HD_Sim", mgvfHD_Sim);
        mobjGlobalVars.AddSingle("HD_Sim_001", mgvfHD_Sim_001);
        mobjGlobalVars.AddSingle("HD_Sim_002", mgvfHD_Sim_002);
        mobjGlobalVars.AddSingle("HD_Sim_003", mgvfHD_Sim_003);
        mobjGlobalVars.AddSingle("HR_Sim", mgvfHR_Sim);
        mobjGlobalVars.AddSingle("Hx_AvgTemp", mgvfHx_AvgTemp);
        mobjGlobalVars.AddSingle("Idle_Sim", mgvfIdle_Sim);
        mobjGlobalVars.AddSingle("MF_Previous", mgvfMF_Previous);
        mobjGlobalVars.AddSingle("MF_Sim_001", mgvfMF_Sim_001);
        mobjGlobalVars.AddSingle("MF_Sim_002", mgvfMF_Sim_002);
        mobjGlobalVars.AddSingle("MF_SpeedDelayed", mgvfMF_SpeedDelayed);
        mobjGlobalVars.AddSingle("MF_SpeedDelayType", mgvfMF_SpeedDelayType);
        mobjGlobalVars.AddSingle("PG_Sim", mgvfPG_Sim);
        mobjGlobalVars.AddSingle("Pstatic", mgvfPstatic);
        mobjGlobalVars.AddSingle("Q_Sim", mgvfQ_Sim);
        mobjGlobalVars.AddSingle("Q_Sim_001", mgvfQ_Sim_001);
        mobjGlobalVars.AddSingle("ResetSim", mgvfResetSim);
        mobjGlobalVars.AddSingle("SF_Sim", mgvfSF_Sim);
        mobjGlobalVars.AddSingle("SR_Sim", mgvfSR_Sim);
        mobjGlobalVars.AddSingle("SS_Sim", mgvfSS_Sim);
        mobjGlobalVars.AddSingle("TimeInterval_Sim", mgvfTimeInterval_Sim);
        mobjGlobalVars.AddSingle("TimeInterval_Sim_001", mgvfTimeInterval_Sim_001);
        mobjGlobalVars.AddSingle("WS_Sim", mgvfWS_Sim);
        mobjGlobalVars.AddSingle("WS_Sim_001", mgvfWS_Sim_001);
        mobjGlobalVars.AddSingle("Airtemp", mopAirtemp);
        mobjGlobalVars.AddSingle("AL1_FaultSts", mopAL1_FaultSts);
        mobjGlobalVars.AddSingle("AL2_FaultSts", mopAL2_FaultSts);
        mobjGlobalVars.AddSingle("BP_Raw", mopBP_Raw);
        mobjGlobalVars.AddSingle("CCE_Perm", mopCCE_Perm);
        mobjGlobalVars.AddSingle("CCE_Remote", mopCCE_Remote);
        mobjGlobalVars.AddSingle("CCE_RunSts", mopCCE_RunSts);
        mobjGlobalVars.AddSingle("CS_BrineSts", mopCS_BrineSts);
        mobjGlobalVars.AddSingle("CS_ChSumFlt", mopCS_ChSumFlt);
        mobjGlobalVars.AddSingle("CS_CritFault", mopCS_CritFault);
        mobjGlobalVars.AddSingle("CS_GenFault", mopCS_GenFault);
        mobjGlobalVars.AddSingle("CS_Health", mopCS_Health);
        mobjGlobalVars.AddSingle("CS_HtrFltSts", mopCS_HtrFltSts);
        mobjGlobalVars.AddSingle("CS_HtrRunSts", mopCS_HtrRunSts);
        mobjGlobalVars.AddSingle("CS_LTclsSts", mopCS_LTclsSts);
        mobjGlobalVars.AddSingle("CS_LTopnSts", mopCS_LTopnSts);
        mobjGlobalVars.AddSingle("CS_MTclsSts", mopCS_MTclsSts);
        mobjGlobalVars.AddSingle("CS_MTopnSts", mopCS_MTopnSts);
        mobjGlobalVars.AddSingle("CS_RunSts", mopCS_RunSts);
        mobjGlobalVars.AddSingle("Ctrl_Power", mopCtrl_Power);
        mobjGlobalVars.AddSingle("CV_Position", mopCV_Position);
        mobjGlobalVars.AddSingle("DCA_FaultSts", mopDCA_FaultSts);
        mobjGlobalVars.AddSingle("DCA_Perm", mopDCA_Perm);
        mobjGlobalVars.AddSingle("DCA_Remote", mopDCA_Remote);
        mobjGlobalVars.AddSingle("DCA_RunSts", mopDCA_RunSts);
        mobjGlobalVars.AddSingle("DCA_Temp", mopDCA_Temp);
        mobjGlobalVars.AddSingle("DeltaP_Hi", mopDeltaP_Hi);
        mobjGlobalVars.AddSingle("DeltaP_Lo", mopDeltaP_Lo);
        mobjGlobalVars.AddSingle("Door_Dy", mopDoor_Dy);
        mobjGlobalVars.AddSingle("Door_UpLg_AS", mopDoor_UpLg_AS);
        mobjGlobalVars.AddSingle("DV_Gly_Temp", mopDV_Gly_Temp);
        mobjGlobalVars.AddSingle("DV_Oil_Temp", mopDV_Oil_Temp);
        mobjGlobalVars.AddSingle("Dy_Estop_Sts", mopDy_Estop_Sts);
        mobjGlobalVars.AddSingle("Dy_Fault", mopDy_Fault);
        mobjGlobalVars.AddSingle("DyPit_Temp", mopDyPit_Temp);
        mobjGlobalVars.AddSingle("EPB_Ctrl_Rm", mopEPB_Ctrl_Rm);
        mobjGlobalVars.AddSingle("EPB_TS_NW", mopEPB_TS_NW);
        mobjGlobalVars.AddSingle("EPB_TS_SW", mopEPB_TS_SW);
        mobjGlobalVars.AddSingle("EPB_UpLeg_AS", mopEPB_UpLeg_AS);
        mobjGlobalVars.AddSingle("EStop_SumSts", mopEStop_SumSts);
        mobjGlobalVars.AddSingle("EX_FanSpeed", mopEX_FanSpeed);
        mobjGlobalVars.AddSingle("EX_Fault", mopEX_Fault);
        mobjGlobalVars.AddSingle("EX_Perm", mopEX_Perm);
        mobjGlobalVars.AddSingle("EX_Remote", mopEX_Remote);
        mobjGlobalVars.AddSingle("EX_RunSts", mopEX_RunSts);
        mobjGlobalVars.AddSingle("FaultBus", mopFaultBus);
        mobjGlobalVars.AddSingle("Fire_DYSup", mopFire_DYSup);
        mobjGlobalVars.AddSingle("Fire_Sum", mopFire_Sum);
        mobjGlobalVars.AddSingle("Gas_Sys_Sts", mopGas_Sys_Sts);
        mobjGlobalVars.AddSingle("HD_Mirror", mopHD_Mirror);
        mobjGlobalVars.AddSingle("HD_SmplFlow", mopHD_SmplFlow);
        mobjGlobalVars.AddSingle("HR_Brn_Temp", mopHR_Brn_Temp);
        mobjGlobalVars.AddSingle("HR_EstopSts", mopHR_EstopSts);
        mobjGlobalVars.AddSingle("HR_FaultSts", mopHR_FaultSts);
        mobjGlobalVars.AddSingle("HR_Perm", mopHR_Perm);
        mobjGlobalVars.AddSingle("HR1_FaultSts", mopHR1_FaultSts);
        mobjGlobalVars.AddSingle("HR1_HxInTmp", mopHR1_HxInTmp);
        mobjGlobalVars.AddSingle("HR1_HxOutTmp", mopHR1_HxOutTmp);
        mobjGlobalVars.AddSingle("HR1_PnlTemp", mopHR1_PnlTemp);
        mobjGlobalVars.AddSingle("HR2_FaultSts", mopHR2_FaultSts);
        mobjGlobalVars.AddSingle("HR2_HxInTmp", mopHR2_HxInTmp);
        mobjGlobalVars.AddSingle("HR2_HxOutTmp", mopHR2_HxOutTmp);
        mobjGlobalVars.AddSingle("HR2_PnlTemp", mopHR2_PnlTemp);
        mobjGlobalVars.AddSingle("HR3_FaultSts", mopHR3_FaultSts);
        mobjGlobalVars.AddSingle("HR3_HxInTmp", mopHR3_HxInTmp);
        mobjGlobalVars.AddSingle("HR3_HxOutTmp", mopHR3_HxOutTmp);
        mobjGlobalVars.AddSingle("HR3_PnlTemp", mopHR3_PnlTemp);
        mobjGlobalVars.AddSingle("HR4_FaultSts", mopHR4_FaultSts);
        mobjGlobalVars.AddSingle("HR4_HxInTmp", mopHR4_HxInTmp);
        mobjGlobalVars.AddSingle("HR4_HxOutTmp", mopHR4_HxOutTmp);
        mobjGlobalVars.AddSingle("HR4_PnlTemp", mopHR4_PnlTemp);
        mobjGlobalVars.AddSingle("HX_HiDeltaP", mopHX_HiDeltaP);
        mobjGlobalVars.AddSingle("HxIn_Temp", mopHxIn_Temp);
        mobjGlobalVars.AddSingle("HxOut_Temp", mopHxOut_Temp);
        mobjGlobalVars.AddSingle("IDL_Fault", mopIDL_Fault);
        mobjGlobalVars.AddSingle("IDL_FlCldSts", mopIDL_FlCldSts);
        mobjGlobalVars.AddSingle("IDL_FlOpdSts", mopIDL_FlOpdSts);
        mobjGlobalVars.AddSingle("IDL_ModeOff", mopIDL_ModeOff);
        mobjGlobalVars.AddSingle("IDL_ModeOn", mopIDL_ModeOn);
        mobjGlobalVars.AddSingle("IDL_Perm", mopIDL_Perm);
        mobjGlobalVars.AddSingle("IDL_Remote", mopIDL_Remote);
        mobjGlobalVars.AddSingle("IDL_ShCldSts", mopIDL_ShCldSts);
        mobjGlobalVars.AddSingle("IDL_ShOpdSts", mopIDL_ShOpdSts);
        mobjGlobalVars.AddSingle("Key_Dy_Door", mopKey_Dy_Door);
        mobjGlobalVars.AddSingle("Key_Veh_Door", mopKey_Veh_Door);
        mobjGlobalVars.AddSingle("Lo_Plen_P", mopLo_Plen_P);
        mobjGlobalVars.AddSingle("LoLo_Plen_P", mopLoLo_Plen_P);
        mobjGlobalVars.AddSingle("LTBrn_Temp", mopLTBrn_Temp);
        mobjGlobalVars.AddSingle("MF_Fault", mopMF_Fault);
        mobjGlobalVars.AddSingle("MF_Perm", mopMF_Perm);
        mobjGlobalVars.AddSingle("MF_Remote", mopMF_Remote);
        mobjGlobalVars.AddSingle("MF_RunSts", mopMF_RunSts);
        mobjGlobalVars.AddSingle("MF_SpdLim1", mopMF_SpdLim1);
        mobjGlobalVars.AddSingle("MF_SpdLim2", mopMF_SpdLim2);
        mobjGlobalVars.AddSingle("MF_Speed", mopMF_Speed);
        mobjGlobalVars.AddSingle("MF_SupplySts", mopMF_SupplySts);
        mobjGlobalVars.AddSingle("MF_Warning", mopMF_Warning);
        mobjGlobalVars.AddSingle("MixBrn_Temp", mopMixBrn_Temp);
        mobjGlobalVars.AddSingle("MTBrn_Temp", mopMTBrn_Temp);
        mobjGlobalVars.AddSingle("MUA_DP", mopMUA_DP);
        mobjGlobalVars.AddSingle("MUA_Flow", mopMUA_Flow);
        mobjGlobalVars.AddSingle("MUA_HX109Flt", mopMUA_HX109Flt);
        mobjGlobalVars.AddSingle("MUA_Temp", mopMUA_Temp);
        mobjGlobalVars.AddSingle("P201_FltSts", mopP201_FltSts);
        mobjGlobalVars.AddSingle("P201_RemSts", mopP201_RemSts);
        mobjGlobalVars.AddSingle("P201_RunSts", mopP201_RunSts);
        mobjGlobalVars.AddSingle("Pabs", mopPabs);
        mobjGlobalVars.AddSingle("PG_FanFault", mopPG_FanFault);
        mobjGlobalVars.AddSingle("PG_FanPerm", mopPG_FanPerm);
        mobjGlobalVars.AddSingle("PG_FanRemote", mopPG_FanRemote);
        mobjGlobalVars.AddSingle("PG_FanRunH", mopPG_FanRunH);
        mobjGlobalVars.AddSingle("PG_FanRunL", mopPG_FanRunL);
        mobjGlobalVars.AddSingle("PG_HL_Tmr", mopPG_HL_Tmr);
        mobjGlobalVars.AddSingle("PG_HLActive", mopPG_HLActive);
        mobjGlobalVars.AddSingle("PG_HLRunSts", mopPG_HLRunSts);
        mobjGlobalVars.AddSingle("PGDy_FanFlt", mopPGDy_FanFlt);
        mobjGlobalVars.AddSingle("PGDy_FanPerm", mopPGDy_FanPerm);
        mobjGlobalVars.AddSingle("PGDy_FRunH", mopPGDy_FRunH);
        mobjGlobalVars.AddSingle("PGDy_FRunL", mopPGDy_FRunL);
        mobjGlobalVars.AddSingle("PGDy_HL_Tmr", mopPGDy_HL_Tmr);
        mobjGlobalVars.AddSingle("PGDy_HLRnSts", mopPGDy_HLRnSts);
        mobjGlobalVars.AddSingle("PGDy_Remote", mopPGDy_Remote);
        mobjGlobalVars.AddSingle("RelHumidity", mopRelHumidity);
        mobjGlobalVars.AddSingle("RV_Position", mopRV_Position);
        mobjGlobalVars.AddSingle("SR1_AirTemp", mopSR1_AirTemp);
        mobjGlobalVars.AddSingle("SR1_FaultSts", mopSR1_FaultSts);
        mobjGlobalVars.AddSingle("SR1_Master", mopSR1_Master);
        mobjGlobalVars.AddSingle("SR1_RunSts", mopSR1_RunSts);
        mobjGlobalVars.AddSingle("SR1_SetptFbk", mopSR1_SetptFbk);
        mobjGlobalVars.AddSingle("SR2_AirTemp", mopSR2_AirTemp);
        mobjGlobalVars.AddSingle("SR2_FaultSts", mopSR2_FaultSts);
        mobjGlobalVars.AddSingle("SR2_Master", mopSR2_Master);
        mobjGlobalVars.AddSingle("SR2_RunSts", mopSR2_RunSts);
        mobjGlobalVars.AddSingle("SR2_SetptFbk", mopSR2_SetptFbk);
        mobjGlobalVars.AddSingle("SR3_AirTemp", mopSR3_AirTemp);
        mobjGlobalVars.AddSingle("SR3_FaultSts", mopSR3_FaultSts);
        mobjGlobalVars.AddSingle("SR3_Master", mopSR3_Master);
        mobjGlobalVars.AddSingle("SR3_RunSts", mopSR3_RunSts);
        mobjGlobalVars.AddSingle("SR3_SetptFbk", mopSR3_SetptFbk);
        mobjGlobalVars.AddSingle("SR4_AirTemp", mopSR4_AirTemp);
        mobjGlobalVars.AddSingle("SR4_FaultSts", mopSR4_FaultSts);
        mobjGlobalVars.AddSingle("SR4_Master", mopSR4_Master);
        mobjGlobalVars.AddSingle("SR4_RunSts", mopSR4_RunSts);
        mobjGlobalVars.AddSingle("SR4_SetptFbk", mopSR4_SetptFbk);
        mobjGlobalVars.AddSingle("SR5_AirTemp", mopSR5_AirTemp);
        mobjGlobalVars.AddSingle("SR5_FaultSts", mopSR5_FaultSts);
        mobjGlobalVars.AddSingle("SR5_Master", mopSR5_Master);
        mobjGlobalVars.AddSingle("SR5_RunSts", mopSR5_RunSts);
        mobjGlobalVars.AddSingle("SR5_SetptFbk", mopSR5_SetptFbk);
        mobjGlobalVars.AddSingle("SR6_AirTemp", mopSR6_AirTemp);
        mobjGlobalVars.AddSingle("SR6_FaultSts", mopSR6_FaultSts);
        mobjGlobalVars.AddSingle("SR6_Master", mopSR6_Master);
        mobjGlobalVars.AddSingle("SR6_RunSts", mopSR6_RunSts);
        mobjGlobalVars.AddSingle("SR6_SetptFbk", mopSR6_SetptFbk);
        mobjGlobalVars.AddSingle("SS_ArmExtSts", mopSS_ArmExtSts);
        mobjGlobalVars.AddSingle("SS_ArmRetSts", mopSS_ArmRetSts);
        mobjGlobalVars.AddSingle("SS_Fan1Fault", mopSS_Fan1Fault);
        mobjGlobalVars.AddSingle("SS_Fan1Rem", mopSS_Fan1Rem);
        mobjGlobalVars.AddSingle("SS_Fan2Fault", mopSS_Fan2Fault);
        mobjGlobalVars.AddSingle("SS_Fan2Rem", mopSS_Fan2Rem);
        mobjGlobalVars.AddSingle("SS_FanPerm", mopSS_FanPerm);
        mobjGlobalVars.AddSingle("SS_Fn1RunSts", mopSS_Fn1RunSts);
        mobjGlobalVars.AddSingle("SS_Fn2RunSts", mopSS_Fn2RunSts);
        mobjGlobalVars.AddSingle("SS_HiTemp", mopSS_HiTemp);
        mobjGlobalVars.AddSingle("SS_InputPyro", mopSS_InputPyro);
        mobjGlobalVars.AddSingle("SS_Perm", mopSS_Perm);
        mobjGlobalVars.AddSingle("SS_RunSts", mopSS_RunSts);
        mobjGlobalVars.AddSingle("SteamPress", mopSteamPress);
        mobjGlobalVars.AddSingle("SV_Position", mopSV_Position);
        mobjGlobalVars.AddSingle("TA1_FaultSts", mopTA1_FaultSts);
        mobjGlobalVars.AddSingle("TA2_FaultSts", mopTA2_FaultSts);
        mobjGlobalVars.AddSingle("VH_IgnPerm", mopVH_IgnPerm);
        mobjGlobalVars.AddSingle("VH_IgnSts", mopVH_IgnSts);
        // On Error GoTo Error — nested handler not restructured
        mobjStatuses = new Statuses();
        mobjStatuses.UseSPG(SPG_STATUSES);
        mobjStatuses.Add(STS_SIM_INITIALIZED, false, "Sim Not Initialized", "Sim Initialized");
        objProcess = new Process();
        mbReady = true;
        WaitForProcessesStartupComplete();
        mlID = AddTalentProcessGetID(objProcess.Name.ToString(), objProcess);
        // On Error Resume Next — nested handler; see enclosing try/catch
        // On Error GoTo Error — nested handler not restructured
        // On Error GoTo Error_Configurations — nested handler not restructured
        sClassName = "ParManager";
        mobjParManager = new ParManager();
        mobjParManager.OwnerID = mlID;
        mobjParManager.AutoRefreshEnabled = true;
        mobjParManager.IgnoreLoadedEvents = true;
        mobjParManager.IgnoreUnloadedEvents = true;
        mobjParManager.AddWithoutFlag("AirDensity", mipAirDensity, ioInput);
        mobjParManager.AddWithoutFlag("AT_Rate", mipAT_Rate, ioInput);
        mobjParManager.AddWithoutFlag("CCE_Start", mipCCE_Start, ioInput);
        mobjParManager.AddWithoutFlag("CS_BrineReq", mipCS_BrineReq, ioInput);
        mobjParManager.AddWithoutFlag("CS_HtrSetpt", mipCS_HtrSetpt, ioInput);
        mobjParManager.AddWithoutFlag("CV_Setpt", mipCV_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("DCA_Start", mipDCA_Start, ioInput);
        mobjParManager.AddWithoutFlag("DCA_Target", mipDCA_Target, ioInput);
        mobjParManager.AddWithoutFlag("Dewpoint", mipDewpoint, ioInput);
        mobjParManager.AddWithoutFlag("DP_Target", mipDP_Target, ioInput);
        mobjParManager.AddWithoutFlag("DV_Gly_Start", mipDV_Gly_Start, ioInput);
        mobjParManager.AddWithoutFlag("DV_Gly_TTgt", mipDV_Gly_TTgt, ioInput);
        mobjParManager.AddWithoutFlag("DV_Oil_Start", mipDV_Oil_Start, ioInput);
        mobjParManager.AddWithoutFlag("DV_Oil_TTgt", mipDV_Oil_TTgt, ioInput);
        mobjParManager.AddWithoutFlag("Dy_CommSts", mipDy_CommSts, ioInput);
        mobjParManager.AddWithoutFlag("Dy_Remote", mipDy_Remote, ioInput);
        mobjParManager.AddWithoutFlag("DyPit_Start", mipDyPit_Start, ioInput);
        mobjParManager.AddWithoutFlag("DyPit_Target", mipDyPit_Target, ioInput);
        mobjParManager.AddWithoutFlag("EX_FanSetpt", mipEX_FanSetpt, ioInput);
        mobjParManager.AddWithoutFlag("EX_Start", mipEX_Start, ioInput);
        mobjParManager.AddWithoutFlag("HR_BrnSetpt", mipHR_BrnSetpt, ioInput);
        mobjParManager.AddWithoutFlag("HR_RunSts", mipHR_RunSts, ioInput);
        mobjParManager.AddWithoutFlag("HR1_CVSetpt", mipHR1_CVSetpt, ioInput);
        mobjParManager.AddWithoutFlag("HR1_HtrSetpt", mipHR1_HtrSetpt, ioInput);
        mobjParManager.AddWithoutFlag("HR1_RunSts", mipHR1_RunSts, ioInput);
        mobjParManager.AddWithoutFlag("HR2_CVSetpt", mipHR2_CVSetpt, ioInput);
        mobjParManager.AddWithoutFlag("HR2_HtrSetpt", mipHR2_HtrSetpt, ioInput);
        mobjParManager.AddWithoutFlag("HR2_RunSts", mipHR2_RunSts, ioInput);
        mobjParManager.AddWithoutFlag("HR3_CVSetpt", mipHR3_CVSetpt, ioInput);
        mobjParManager.AddWithoutFlag("HR3_HtrSetpt", mipHR3_HtrSetpt, ioInput);
        mobjParManager.AddWithoutFlag("HR3_RunSts", mipHR3_RunSts, ioInput);
        mobjParManager.AddWithoutFlag("HR4_CVSetpt", mipHR4_CVSetpt, ioInput);
        mobjParManager.AddWithoutFlag("HR4_HtrSetpt", mipHR4_HtrSetpt, ioInput);
        mobjParManager.AddWithoutFlag("HR4_RunSts", mipHR4_RunSts, ioInput);
        mobjParManager.AddWithoutFlag("IDL_Start", mipIDL_Start, ioInput);
        mobjParManager.AddWithoutFlag("LTBrn_Warn", mipLTBrn_Warn, ioInput);
        mobjParManager.AddWithoutFlag("MF_Setpt", mipMF_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("MF_Start", mipMF_Start, ioInput);
        mobjParManager.AddWithoutFlag("MTBrn_Warn", mipMTBrn_Warn, ioInput);
        mobjParManager.AddWithoutFlag("P201_RunReq", mipP201_RunReq, ioInput);
        mobjParManager.AddWithoutFlag("PG_HL_DelRst", mipPG_HL_DelRst, ioInput);
        mobjParManager.AddWithoutFlag("PG_HL_Start", mipPG_HL_Start, ioInput);
        mobjParManager.AddWithoutFlag("PG_LL_Start", mipPG_LL_Start, ioInput);
        mobjParManager.AddWithoutFlag("PGDy_HL_DRst", mipPGDy_HL_DRst, ioInput);
        mobjParManager.AddWithoutFlag("PGDy_HL_Strt", mipPGDy_HL_Strt, ioInput);
        mobjParManager.AddWithoutFlag("PGDy_LL_Strt", mipPGDy_LL_Strt, ioInput);
        mobjParManager.AddWithoutFlag("PStatic", mipPStatic, ioInput);
        mobjParManager.AddWithoutFlag("Q", mipQ, ioInput);
        mobjParManager.AddWithoutFlag("RV_Setpt", mipRV_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("SR1_RunReq", mipSR1_RunReq, ioInput);
        mobjParManager.AddWithoutFlag("SR1_Setpt", mipSR1_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("SR2_RunReq", mipSR2_RunReq, ioInput);
        mobjParManager.AddWithoutFlag("SR2_Setpt", mipSR2_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("SR3_RunReq", mipSR3_RunReq, ioInput);
        mobjParManager.AddWithoutFlag("SR3_Setpt", mipSR3_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("SR4_RunReq", mipSR4_RunReq, ioInput);
        mobjParManager.AddWithoutFlag("SR4_Setpt", mipSR4_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("SR5_RunReq", mipSR5_RunReq, ioInput);
        mobjParManager.AddWithoutFlag("SR5_Setpt", mipSR5_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("SR6_RunReq", mipSR6_RunReq, ioInput);
        mobjParManager.AddWithoutFlag("SR6_Setpt", mipSR6_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("SS_ArmExtReq", mipSS_ArmExtReq, ioInput);
        mobjParManager.AddWithoutFlag("SS_FanStart", mipSS_FanStart, ioInput);
        mobjParManager.AddWithoutFlag("SS_InvCoef0", mipSS_InvCoef0, ioInput);
        mobjParManager.AddWithoutFlag("SS_InvCoef1", mipSS_InvCoef1, ioInput);
        mobjParManager.AddWithoutFlag("SS_InvCoef2", mipSS_InvCoef2, ioInput);
        mobjParManager.AddWithoutFlag("SS_InvCoef3", mipSS_InvCoef3, ioInput);
        mobjParManager.AddWithoutFlag("SS_InvCoef4", mipSS_InvCoef4, ioInput);
        mobjParManager.AddWithoutFlag("SS_Power", mipSS_Power, ioInput);
        mobjParManager.AddWithoutFlag("SS_RunReq", mipSS_RunReq, ioInput);
        mobjParManager.AddWithoutFlag("SV_Setpt", mipSV_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("VH_IgnEnable", mipVH_IgnEnable, ioInput);
        mobjParManager.AddWithoutFlag("WS_Setpt", mipWS_Setpt, ioInput);
        mobjParManager.AddWithoutFlag("Airtemp", mopAirtemp, ioOutput);
        mobjParManager.AddWithoutFlag("AL1_FaultSts", mopAL1_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("AL2_FaultSts", mopAL2_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("BP_Raw", mopBP_Raw, ioOutput);
        mobjParManager.AddWithoutFlag("CCE_Perm", mopCCE_Perm, ioOutput);
        mobjParManager.AddWithoutFlag("CCE_Remote", mopCCE_Remote, ioOutput);
        mobjParManager.AddWithoutFlag("CCE_RunSts", mopCCE_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("CS_BrineSts", mopCS_BrineSts, ioOutput);
        mobjParManager.AddWithoutFlag("CS_ChSumFlt", mopCS_ChSumFlt, ioOutput);
        mobjParManager.AddWithoutFlag("CS_CritFault", mopCS_CritFault, ioOutput);
        mobjParManager.AddWithoutFlag("CS_GenFault", mopCS_GenFault, ioOutput);
        mobjParManager.AddWithoutFlag("CS_Health", mopCS_Health, ioOutput);
        mobjParManager.AddWithoutFlag("CS_HtrFltSts", mopCS_HtrFltSts, ioOutput);
        mobjParManager.AddWithoutFlag("CS_HtrRunSts", mopCS_HtrRunSts, ioOutput);
        mobjParManager.AddWithoutFlag("CS_LTclsSts", mopCS_LTclsSts, ioOutput);
        mobjParManager.AddWithoutFlag("CS_LTopnSts", mopCS_LTopnSts, ioOutput);
        mobjParManager.AddWithoutFlag("CS_MTclsSts", mopCS_MTclsSts, ioOutput);
        mobjParManager.AddWithoutFlag("CS_MTopnSts", mopCS_MTopnSts, ioOutput);
        mobjParManager.AddWithoutFlag("CS_RunSts", mopCS_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("Ctrl_Power", mopCtrl_Power, ioOutput);
        mobjParManager.AddWithoutFlag("CV_Position", mopCV_Position, ioOutput);
        mobjParManager.AddWithoutFlag("DCA_FaultSts", mopDCA_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("DCA_Perm", mopDCA_Perm, ioOutput);
        mobjParManager.AddWithoutFlag("DCA_Remote", mopDCA_Remote, ioOutput);
        mobjParManager.AddWithoutFlag("DCA_RunSts", mopDCA_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("DCA_Temp", mopDCA_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("DeltaP_Hi", mopDeltaP_Hi, ioOutput);
        mobjParManager.AddWithoutFlag("DeltaP_Lo", mopDeltaP_Lo, ioOutput);
        mobjParManager.AddWithoutFlag("Door_Dy", mopDoor_Dy, ioOutput);
        mobjParManager.AddWithoutFlag("Door_UpLg_AS", mopDoor_UpLg_AS, ioOutput);
        mobjParManager.AddWithoutFlag("DV_Gly_Temp", mopDV_Gly_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("DV_Oil_Temp", mopDV_Oil_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("Dy_Estop_Sts", mopDy_Estop_Sts, ioOutput);
        mobjParManager.AddWithoutFlag("Dy_Fault", mopDy_Fault, ioOutput);
        mobjParManager.AddWithoutFlag("DyPit_Temp", mopDyPit_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("EPB_Ctrl_Rm", mopEPB_Ctrl_Rm, ioOutput);
        mobjParManager.AddWithoutFlag("EPB_TS_NW", mopEPB_TS_NW, ioOutput);
        mobjParManager.AddWithoutFlag("EPB_TS_SW", mopEPB_TS_SW, ioOutput);
        mobjParManager.AddWithoutFlag("EPB_UpLeg_AS", mopEPB_UpLeg_AS, ioOutput);
        mobjParManager.AddWithoutFlag("EStop_SumSts", mopEStop_SumSts, ioOutput);
        mobjParManager.AddWithoutFlag("EX_FanSpeed", mopEX_FanSpeed, ioOutput);
        mobjParManager.AddWithoutFlag("EX_Fault", mopEX_Fault, ioOutput);
        mobjParManager.AddWithoutFlag("EX_Perm", mopEX_Perm, ioOutput);
        mobjParManager.AddWithoutFlag("EX_Remote", mopEX_Remote, ioOutput);
        mobjParManager.AddWithoutFlag("EX_RunSts", mopEX_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("FaultBus", mopFaultBus, ioOutput);
        mobjParManager.AddWithoutFlag("Fire_DYSup", mopFire_DYSup, ioOutput);
        mobjParManager.AddWithoutFlag("Fire_Sum", mopFire_Sum, ioOutput);
        mobjParManager.AddWithoutFlag("Gas_Sys_Sts", mopGas_Sys_Sts, ioOutput);
        mobjParManager.AddWithoutFlag("HD_Mirror", mopHD_Mirror, ioOutput);
        mobjParManager.AddWithoutFlag("HD_SmplFlow", mopHD_SmplFlow, ioOutput);
        mobjParManager.AddWithoutFlag("HR_Brn_Temp", mopHR_Brn_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("HR_EstopSts", mopHR_EstopSts, ioOutput);
        mobjParManager.AddWithoutFlag("HR_FaultSts", mopHR_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("HR_Perm", mopHR_Perm, ioOutput);
        mobjParManager.AddWithoutFlag("HR1_FaultSts", mopHR1_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("HR1_HxInTmp", mopHR1_HxInTmp, ioOutput);
        mobjParManager.AddWithoutFlag("HR1_HxOutTmp", mopHR1_HxOutTmp, ioOutput);
        mobjParManager.AddWithoutFlag("HR1_PnlTemp", mopHR1_PnlTemp, ioOutput);
        mobjParManager.AddWithoutFlag("HR2_FaultSts", mopHR2_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("HR2_HxInTmp", mopHR2_HxInTmp, ioOutput);
        mobjParManager.AddWithoutFlag("HR2_HxOutTmp", mopHR2_HxOutTmp, ioOutput);
        mobjParManager.AddWithoutFlag("HR2_PnlTemp", mopHR2_PnlTemp, ioOutput);
        mobjParManager.AddWithoutFlag("HR3_FaultSts", mopHR3_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("HR3_HxInTmp", mopHR3_HxInTmp, ioOutput);
        mobjParManager.AddWithoutFlag("HR3_HxOutTmp", mopHR3_HxOutTmp, ioOutput);
        mobjParManager.AddWithoutFlag("HR3_PnlTemp", mopHR3_PnlTemp, ioOutput);
        mobjParManager.AddWithoutFlag("HR4_FaultSts", mopHR4_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("HR4_HxInTmp", mopHR4_HxInTmp, ioOutput);
        mobjParManager.AddWithoutFlag("HR4_HxOutTmp", mopHR4_HxOutTmp, ioOutput);
        mobjParManager.AddWithoutFlag("HR4_PnlTemp", mopHR4_PnlTemp, ioOutput);
        mobjParManager.AddWithoutFlag("HX_HiDeltaP", mopHX_HiDeltaP, ioOutput);
        mobjParManager.AddWithoutFlag("HxIn_Temp", mopHxIn_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("HxOut_Temp", mopHxOut_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("IDL_Fault", mopIDL_Fault, ioOutput);
        mobjParManager.AddWithoutFlag("IDL_FlCldSts", mopIDL_FlCldSts, ioOutput);
        mobjParManager.AddWithoutFlag("IDL_FlOpdSts", mopIDL_FlOpdSts, ioOutput);
        mobjParManager.AddWithoutFlag("IDL_ModeOff", mopIDL_ModeOff, ioOutput);
        mobjParManager.AddWithoutFlag("IDL_ModeOn", mopIDL_ModeOn, ioOutput);
        mobjParManager.AddWithoutFlag("IDL_Perm", mopIDL_Perm, ioOutput);
        mobjParManager.AddWithoutFlag("IDL_Remote", mopIDL_Remote, ioOutput);
        mobjParManager.AddWithoutFlag("IDL_ShCldSts", mopIDL_ShCldSts, ioOutput);
        mobjParManager.AddWithoutFlag("IDL_ShOpdSts", mopIDL_ShOpdSts, ioOutput);
        mobjParManager.AddWithoutFlag("Key_Dy_Door", mopKey_Dy_Door, ioOutput);
        mobjParManager.AddWithoutFlag("Key_Veh_Door", mopKey_Veh_Door, ioOutput);
        mobjParManager.AddWithoutFlag("Lo_Plen_P", mopLo_Plen_P, ioOutput);
        mobjParManager.AddWithoutFlag("LoLo_Plen_P", mopLoLo_Plen_P, ioOutput);
        mobjParManager.AddWithoutFlag("LTBrn_Temp", mopLTBrn_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("MF_Fault", mopMF_Fault, ioOutput);
        mobjParManager.AddWithoutFlag("MF_Perm", mopMF_Perm, ioOutput);
        mobjParManager.AddWithoutFlag("MF_Remote", mopMF_Remote, ioOutput);
        mobjParManager.AddWithoutFlag("MF_RunSts", mopMF_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("MF_SpdLim1", mopMF_SpdLim1, ioOutput);
        mobjParManager.AddWithoutFlag("MF_SpdLim2", mopMF_SpdLim2, ioOutput);
        mobjParManager.AddWithoutFlag("MF_Speed", mopMF_Speed, ioOutput);
        mobjParManager.AddWithoutFlag("MF_SupplySts", mopMF_SupplySts, ioOutput);
        mobjParManager.AddWithoutFlag("MF_Warning", mopMF_Warning, ioOutput);
        mobjParManager.AddWithoutFlag("MixBrn_Temp", mopMixBrn_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("MTBrn_Temp", mopMTBrn_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("MUA_DP", mopMUA_DP, ioOutput);
        mobjParManager.AddWithoutFlag("MUA_Flow", mopMUA_Flow, ioOutput);
        mobjParManager.AddWithoutFlag("MUA_HX109Flt", mopMUA_HX109Flt, ioOutput);
        mobjParManager.AddWithoutFlag("MUA_Temp", mopMUA_Temp, ioOutput);
        mobjParManager.AddWithoutFlag("P201_FltSts", mopP201_FltSts, ioOutput);
        mobjParManager.AddWithoutFlag("P201_RemSts", mopP201_RemSts, ioOutput);
        mobjParManager.AddWithoutFlag("P201_RunSts", mopP201_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("Pabs", mopPabs, ioOutput);
        mobjParManager.AddWithoutFlag("PG_FanFault", mopPG_FanFault, ioOutput);
        mobjParManager.AddWithoutFlag("PG_FanPerm", mopPG_FanPerm, ioOutput);
        mobjParManager.AddWithoutFlag("PG_FanRemote", mopPG_FanRemote, ioOutput);
        mobjParManager.AddWithoutFlag("PG_FanRunH", mopPG_FanRunH, ioOutput);
        mobjParManager.AddWithoutFlag("PG_FanRunL", mopPG_FanRunL, ioOutput);
        mobjParManager.AddWithoutFlag("PG_HL_Tmr", mopPG_HL_Tmr, ioOutput);
        mobjParManager.AddWithoutFlag("PG_HLActive", mopPG_HLActive, ioOutput);
        mobjParManager.AddWithoutFlag("PG_HLRunSts", mopPG_HLRunSts, ioOutput);
        mobjParManager.AddWithoutFlag("PGDy_FanFlt", mopPGDy_FanFlt, ioOutput);
        mobjParManager.AddWithoutFlag("PGDy_FanPerm", mopPGDy_FanPerm, ioOutput);
        mobjParManager.AddWithoutFlag("PGDy_FRunH", mopPGDy_FRunH, ioOutput);
        mobjParManager.AddWithoutFlag("PGDy_FRunL", mopPGDy_FRunL, ioOutput);
        mobjParManager.AddWithoutFlag("PGDy_HL_Tmr", mopPGDy_HL_Tmr, ioOutput);
        mobjParManager.AddWithoutFlag("PGDy_HLRnSts", mopPGDy_HLRnSts, ioOutput);
        mobjParManager.AddWithoutFlag("PGDy_Remote", mopPGDy_Remote, ioOutput);
        mobjParManager.AddWithoutFlag("RelHumidity", mopRelHumidity, ioOutput);
        mobjParManager.AddWithoutFlag("RV_Position", mopRV_Position, ioOutput);
        mobjParManager.AddWithoutFlag("SR1_AirTemp", mopSR1_AirTemp, ioOutput);
        mobjParManager.AddWithoutFlag("SR1_FaultSts", mopSR1_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR1_Master", mopSR1_Master, ioOutput);
        mobjParManager.AddWithoutFlag("SR1_RunSts", mopSR1_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR1_SetptFbk", mopSR1_SetptFbk, ioOutput);
        mobjParManager.AddWithoutFlag("SR2_AirTemp", mopSR2_AirTemp, ioOutput);
        mobjParManager.AddWithoutFlag("SR2_FaultSts", mopSR2_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR2_Master", mopSR2_Master, ioOutput);
        mobjParManager.AddWithoutFlag("SR2_RunSts", mopSR2_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR2_SetptFbk", mopSR2_SetptFbk, ioOutput);
        mobjParManager.AddWithoutFlag("SR3_AirTemp", mopSR3_AirTemp, ioOutput);
        mobjParManager.AddWithoutFlag("SR3_FaultSts", mopSR3_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR3_Master", mopSR3_Master, ioOutput);
        mobjParManager.AddWithoutFlag("SR3_RunSts", mopSR3_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR3_SetptFbk", mopSR3_SetptFbk, ioOutput);
        mobjParManager.AddWithoutFlag("SR4_AirTemp", mopSR4_AirTemp, ioOutput);
        mobjParManager.AddWithoutFlag("SR4_FaultSts", mopSR4_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR4_Master", mopSR4_Master, ioOutput);
        mobjParManager.AddWithoutFlag("SR4_RunSts", mopSR4_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR4_SetptFbk", mopSR4_SetptFbk, ioOutput);
        mobjParManager.AddWithoutFlag("SR5_AirTemp", mopSR5_AirTemp, ioOutput);
        mobjParManager.AddWithoutFlag("SR5_FaultSts", mopSR5_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR5_Master", mopSR5_Master, ioOutput);
        mobjParManager.AddWithoutFlag("SR5_RunSts", mopSR5_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR5_SetptFbk", mopSR5_SetptFbk, ioOutput);
        mobjParManager.AddWithoutFlag("SR6_AirTemp", mopSR6_AirTemp, ioOutput);
        mobjParManager.AddWithoutFlag("SR6_FaultSts", mopSR6_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR6_Master", mopSR6_Master, ioOutput);
        mobjParManager.AddWithoutFlag("SR6_RunSts", mopSR6_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("SR6_SetptFbk", mopSR6_SetptFbk, ioOutput);
        mobjParManager.AddWithoutFlag("SS_ArmExtSts", mopSS_ArmExtSts, ioOutput);
        mobjParManager.AddWithoutFlag("SS_ArmRetSts", mopSS_ArmRetSts, ioOutput);
        mobjParManager.AddWithoutFlag("SS_Fan1Fault", mopSS_Fan1Fault, ioOutput);
        mobjParManager.AddWithoutFlag("SS_Fan1Rem", mopSS_Fan1Rem, ioOutput);
        mobjParManager.AddWithoutFlag("SS_Fan2Fault", mopSS_Fan2Fault, ioOutput);
        mobjParManager.AddWithoutFlag("SS_Fan2Rem", mopSS_Fan2Rem, ioOutput);
        mobjParManager.AddWithoutFlag("SS_FanPerm", mopSS_FanPerm, ioOutput);
        mobjParManager.AddWithoutFlag("SS_Fn1RunSts", mopSS_Fn1RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("SS_Fn2RunSts", mopSS_Fn2RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("SS_HiTemp", mopSS_HiTemp, ioOutput);
        mobjParManager.AddWithoutFlag("SS_InputPyro", mopSS_InputPyro, ioOutput);
        mobjParManager.AddWithoutFlag("SS_Perm", mopSS_Perm, ioOutput);
        mobjParManager.AddWithoutFlag("SS_RunSts", mopSS_RunSts, ioOutput);
        mobjParManager.AddWithoutFlag("SteamPress", mopSteamPress, ioOutput);
        mobjParManager.AddWithoutFlag("SV_Position", mopSV_Position, ioOutput);
        mobjParManager.AddWithoutFlag("TA1_FaultSts", mopTA1_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("TA2_FaultSts", mopTA2_FaultSts, ioOutput);
        mobjParManager.AddWithoutFlag("VH_IgnPerm", mopVH_IgnPerm, ioOutput);
        mobjParManager.AddWithoutFlag("VH_IgnSts", mopVH_IgnSts, ioOutput);
        // On Error GoTo Error — nested handler not restructured
        if (mccfUpdateTime < 0.05)
        {
            mccfUpdateTime = 0.05;
        }

        if (IsWithinIDE())
        {
            mobjForm = new frmControl();
            Load(mobjForm);
            mobjForm.tmrDesignTime.Interval = mccfUpdateTime * 1000;
            mobjForm.tmrDesignTime.Enabled = true;
        }
        else
        {
            tmrMain.Stats.Frequency = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_PROCESSES, null, "TimerFrequency", 10, vbLong);
            tmrMain.Interval = mccfUpdateTime * 1000;
            tmrMain.EventType = TimerPeriodic;
            tmrMain.Notify = this;
            SetProcessPriority(mlDefPriority);
            tmrMain.Enabled = true;
        }

        if (mobjParManager.Configurations.Count > 0)
        {
            mobjCommands.Insert(ocInitialize, "SimInitialize");
        }

        tmrRunOnce = null;
        mobjStats.Reset();
        for (i = 1; i <= 200; i++)
        {
            mobjStats.AddValue((double)mccfUpdateTime);
        }

        tmrElapsed.Reset();
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        EH_MsgBox();
        LogMessage(string.Concat(string.Concat("Error during controller startup: ", sRaiseErrorText), ". See log file for details"), mtFailure);
    Error_CreatingObjects:
        ;
        sRaiseErrorText = string.Concat(string.Concat(string.Concat("Can't create object '", sClassName), "': "), ex.Message);
        bRaiseError = true;
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        EH.Item(1) = sClassName;
        EH_MsgBox();
        goto Done;
    Error_Configurations:
        ;
        EH.VBErr = Err;
        EH.AppError = LoadResString(StdTalentStrings.stsEHCantCreateConfigurations);
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        sRaiseErrorText = EH.Text;
        EH_MsgBox();
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    public Stats Stats
    {
        get
        {
            Stats __result = 0;
            __result = mobjStats;
            return __result;
        }
    }

    private void Suspend()
    {
        Attribute(Suspend.VB_Description == "Suspend control");
        try
        {
            mbInitialized = false;
            mobjStatuses(STS_SIM_INITIALIZED).SetValue(mbInitialized);
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void tmrRunOnce_Timer(int Milliseconds)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        tmrRunOnce.Enabled = false;
        tmrRunOnce = null;
        Start();
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }
}
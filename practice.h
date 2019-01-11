//                                ---- Drive -------   Over   FD       CMFS              A3     ----------------- VOS --------------------   MinFlow g/ccPerDegC
//                                Amp      P      I    shoot  lim       A4      GasFD   limit   Gas Slope  Gas Offst  Liq Slope   Liq Offst  Mult    tempEffect Flags
//CAT_TABLE,Drive Target,!P!,!I!,Overshoot,FD Limit,A4,GasFD,X,!Gas Slope!,!Gas Offst!,!Liq Slope!,!Liq Offst!,Minimum Flow Multiplier,X,X 
const S_CAT   CAT_S_E         = { _3p4, 100.0f,  0.5f, 1.07f,  _14,      _0p0,   _1p0,  _0p425,      _0p0,      _0p0,      _0p0,       _0p0,  50.0f, 0.000015f, 0 };
const S_CAT   CAT_TiTAN1      = { 0.5f, 100.0f,  0.5f, 1.07f, 2.0f,      _0p0,   _1p0, 0.0625f,      _0p0,      _0p0,      _0p0,       _0p0, 100.0f,      _0p0, SSA_TSERIES_COMP };
const S_CAT   CAT_TiTAN2      = { 0.5f,   _150,  _0p7,  _1p1, 2.0f,      _0p0,   _1p0, 0.0625f,      _0p0,      _0p0,      _0p0,       _0p0, 100.0f,      _0p0, SSA_TSERIES_COMP };
//End


#define _3p4     3.4f       // default mV/Hz Target Amp
#define _150   150.0f       // default drive P
#define _0p7     0.7f       // default drive I
#define _1p1     1.1f       // default overshoot
#define _14     14.0f       // default uS FD Limit
#define _1p0     1.0f       // default GasFD
#define _0p425 0.425f       // default mv/Hz A3 limit (.425 is 12.5% of 3.4)
#define _0p0     0.0f       // default for VOS

/*                               3083- 3085- 3087--   3089  3091--  3094--     3159----    3161-- 3163-- 3165--   3089-  3091--  3163- */
//SMV_TABLE,Tone Level,Ramp Time,BL Temp Coeff,Drive SP FCF,Puck P FCF,BL Temp ?Quad? Coeff,dF Tone Spacing,Freq. Drift Limit,Max Sensor Current,X,X 
const S_SMV  SMV_CMF010  = {      5.f, 30.f, 6.970f,  1.7f,  50.f, -0.003533f, 0.2666667f, .067f,   8.f,    0.f,  1.7f, 50000.f, 300.f };
const S_SMV  SMV_CMF025p = {     15.f, 15.f, 7.680f,  3.4f, 250.f, -0.003649f, 0.3333333f, .083f,  30.f,    0.f,  1.7f, 50000.f, 343.f };
const S_SMV  SMV_CMF050  = {     20.f, 15.f, 6.970f,  .85f, 250.f, -0.003533f, 0.3333333f, .083f,  30.f,    0.f,  1.0f, 50000.f, 300.f };
//End

const SDEF   SdefTable[ ] = {
/* CMF       MV                                uS         m2                                 --- Pressure Comp ---  lb/min     kg/sec    %MaxPerC    %      %     kg/m3      */
/*Name-----  Type----- FCF---- K1------ slot.- Mass-- TubeArea-- Category----- SmvCoeffs---  FlowLiq-- Dens-------  Zero Stab- MaxFlow-- TempEff---- liq--- gas-- dens--     */
//COEFF_TABLE,ID String,X,FCF,K1,X,X,TubeID,CAT,S_SMV,PressureEffect_Flow_Liquid,PressureEffect_Density,Zero Stability,X,X,MassFlowAccuracy_Liquid,MassFlowAccuracyMVD_Gas,DensityAccuracy_Liquid */
{"CMF010  ",   CMF010, 0.406f,  9680.f,    s0, 58.0f, 6.587e-6f,     &CAT_S_E,  &SMV_CMF010,  PC_NONE,     PC_NONE, 0.000075f,    0.03f, 0.0001875f, 0.1f,  0.35f,  0.5f }, /* (CMF010M, CMF010N)  */
{"CMF010P ",  CMF010P, 0.689f,  9249.f,    s0, 27.0f, 5.066e-6f,     &CAT_S_E,  &SMV_CMF010,  PC_NONE,     PC_NONE, 0.00015f,     0.03f, 0.0001875f, 0.1f,  0.35f,  2.0f }, /* (CMF010P)           */
{"CMF025  ",   CMF025, 4.339f,  6385.f,    s0, 84.0f, 4.302e-5f,     &CAT_S_E,    &SMV_NONE,  PC_NONE,  0.0000040f, 0.001f,    0.60556f,  0.000125f, 0.1f,  0.35f,  0.5f }, /* (CMF025H, CMF025M)  */
/*Name-----  Type----- FCF---- K1------ slot-- Mass-- TubeArea-- Category----- SmvCoeffs---  FlowLiq-- Dens-------  Zero Stab. MaxFlow-- TempEff---- liq--- gas-- dens--     */
{"CMF025+ ",  CMF025p, 5.000f,  6385.f,    s0, 72.0f, 4.302e-5f,     &CAT_S_E, &SMV_CMF025p,  PC_NONE,  0.0000040f, 0.001f,    0.60556f,  0.000125f, 0.1f,  0.35f,  0.5f }, /* (CMF025H/M/L)       */
{"CMF050  ",   CMF050, 14.87f,  6444.f,    s0, 64.0f, 1.206e-4f,     &CAT_S_E,  &SMV_CMF050,  PC_NONE, -0.0000020f, 0.006f,    1.88999f,  0.000125f, 0.1f,  0.35f,  0.5f }, /* (CMF050H, CMF050M)  */
}; //End

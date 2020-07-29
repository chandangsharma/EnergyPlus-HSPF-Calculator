Attribute VB_Name = "HSPF"
Option Explicit

' MODULE INFORMATION:
' EnergyPlus module Name:  StandardRatings.cc
' Subroutine Name       :  SingelSpeedDXCoolingCoilStandardRatings
'  Author               :  Chandan Sharma

' PURPOSE OF THIS MODULE:
' This module contains the SUBs required to calculate the following standard ratings of HVAC equipment
' Heating Seasonal Performance Factor (HSPF) for Air-Source Direct Expansion Heat Pumps having a single-speed compressor,
'       fixed speed indoor supply air fan
'
' METHODOLOGY EMPLOYED:
' Using the user specified reference capacity, reference COP and performance curves, the DX coil models are executed
' for standard test conditions as specified in ANSI/AHRI 210/240. Then results of the simulated test points
' are processed into standard ratings according to standard's procedures.

' REFERENCES:
' ANSI/AHRI Standard 210/240-2008:  Standard for Performance Rating of Unitary Air-Conditioning and
'                                                       Air-Source Heat Pumps. Arlington, VA:  Air-Conditioning, Heating
'                                                       , and Refrigeration Institute.

' AHRI Standard 210/240-2008 Performance Test Conditions for Unitary Air-to-Air Air-Conditioning and Heat Pump Equipment

Public Const CoolingCoilInletAirWetbulbTempRated As Double = 19.44
Public Const OutdoorUnitInletAirDrybulbTemp As Double = 27.78 ' 27.78C (82F)  Test B (for SEER)
Public Const OutdoorUnitInletAirDrybulbTempRated As Double = 35 ' 35.00C (95F)  Test A (rated capacity)
Public Const AirMassFlowRatioRated As Double = 1 ' AHRI test is at the design flow rate
' and hence AirMassFlowRatio is 1.0
Public Const ConvFromSIToIP As Double = 3.412141633 ' Conversion from SI to IP [3.412 Btu/hr-W]
Public Const DefaultFanPowerPerEvapAirFlowRate As Double = 773.3 ' 365 W/1000 scfm or 773.3 W/(m3/s). The AHRI standard
' specifies a nominal/default fan electric power consumption per rated air
' volume flow rate to account for indoor fan electric power consumption
' when the standard tests are conducted on units that do not have an
' indoor air circulting fan. Used if user doesn't enter a specific value.

' Defrost control  (heat pump only)
'Public Const Timed As Integer = 1 ' defrost cycle is timed
'Public Const OnDemand As Integer = 2 ' defrost cycle occurs only when required
Public Const TotalNumOfStandardDHRs As Integer = 16 ' Total number of standard design heating requirements

Public TotalNumOfTemperatureBins As Variant
' bins for a region

Public Const CorrectionFactor As Double = 0.77 ' A correction factor which tends to improve the agreement
' between calculated and measured building loads, dimensionless.
Public Const CyclicDegradationCoeff As Double = 0.25

Public StandardDesignHeatingRequirement As Variant

Public OutdoorDesignTemperature As Variant

Public OutdoorBinTemperature As Variant

Public RegionOneFracBinHoursAtOutdoorBinTemp As Variant
Public RegionTwoFracBinHoursAtOutdoorBinTemp As Variant
Public RegionThreeFracBinHoursAtOutdoorBinTemp As Variant
Public RegionFourFracBinHoursAtOutdoorBinTemp As Variant
Public RegionFiveFracBinHoursAtOutdoorBinTemp As Variant
Public RegionSixFracBinHoursAtOutdoorBinTemp As Variant

Public Const HeatingIndoorCoilInletAirDBTempRated As Double = 21.11 ' Heating coil entering air dry-bulb temperature in
' degrees C (70F) Test H1, H2 and H3
' (low and High Speed) Std. AHRI 210/240
Public Const HeatingOutdoorCoilInletAirDBTempH0Test As Double = 16.67 ' Outdoor air dry-bulb temp in degrees C (47F)
' Test H0 (low and High Speed) Std. AHRI 210/240
Public Const HeatingOutdoorCoilInletAirDBTempRated As Double = 8.33 ' Outdoor air dry-bulb temp in degrees C (47F)
' Test H1 or rated (low and High Speed) Std. AHRI 210/240
Public Const HeatingOutdoorCoilInletAirDBTempH2Test As Double = 1.67 ' Outdoor air dry-bulb temp in degrees C (35F)
' Test H2 (low and High Speed) Std. AHRI 210/240
Public Const HeatingOutdoorCoilInletAirDBTempH3Test As Double = -8.33 ' Outdoor air dry-bulb temp in degrees C (17F)
' Test H3 (low and High Speed) Std. AHRI 210/240

Public Const FirstHPRowNum As Integer = 6
Public Const FirstHPColumnNum As Integer = 2

Public Const CapFT As Integer = 1
Public Const EIRFT As Integer = 2
Public Const CapFFF As Integer = 3
Public Const EIRFFF As Integer = 4

Public Const CoolingCapacity_ColumnNum As Integer = 3
Public Const RatedTotalHeatingCapacity_ColumnNum As Integer = 4
Public Const RatedAirFlowRate_ColumnNum As Integer = 5
Public Const EvapFanPowerPerVolFlowRate_ColumnNum As Integer = 6
Public Const MinOATCompressor_ColumnNum As Integer = 7
Public Const OATempCompressorOn_ColumnNum As Integer = 8
Public Const COP_ColumnNum As Integer = 9
Public Const CalculatedHSPF_ColumnNum As Integer = 10
Public Const DesiredHSPF_ColumnNum As Integer = 11

Public Const CapFT_C1_ColumnNum As Integer = 14
Public Const EIRFT_C1_ColumnNum As Integer = 24
Public Const CapFFF_C1_ColumnNum As Integer = 34
Public Const EIRFFF_C1_ColumnNum As Integer = 40
Public RegionNum As Integer 'RegionNumber
Public RowNum As Integer
Public DefrostControl As String
Public LastHP As Boolean

Public CapFTempCurveType As String
Public EIRFTempCurveType As String
Public CapFFlowCurveType As String
Public EIRFFlowCurveType As String

Sub SetGlobalConstantArrays()

StandardDesignHeatingRequirement = Array(1465.36, 2930.71, 4396.07, 5861.42, _
7326.78, 8792.14, 10257.49, 11722.85, 14653.56, 17584.27, 20514.98, 23445.7, _
26376.41, 29307.12, 32237.83, 38099.26)

' Standardized DHRs from ANSI/AHRI 210/240
TotalNumOfTemperatureBins = Array(9, 10, 13, 15, 18, 9) ' Total number of temperature

OutdoorDesignTemperature = Array(2.78, -2.78, -8.33, -15#, -23.33, -1.11)
' Outdoor design temperature for a region from ANSI/AHRI 210/240

OutdoorBinTemperature = Array(16.67, 13.89, 11.11, 8.33, 5.56, 2.78, 0#, _
-2.78, -5.56, -8.33, -11.11, -13.89, -16.67, -19.44, -22.22, -25#, -27.78, -30.56)
' Fractional bin hours for different bin temperatures for region one, from ANSI/AHRI 210/240

RegionOneFracBinHoursAtOutdoorBinTemp = Array(0.291, 0.239, 0.194, 0.129, 0.081, 0.041, _
0.019, 0.005, 0.001, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
' Fractional bin hours for different bin temperatures for region two, from ANSI/AHRI 210/240

RegionTwoFracBinHoursAtOutdoorBinTemp = Array(0.215, 0.189, 0.163, 0.143, 0.112, 0.088, _
0.056, 0.024, 0.008, 0.002, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
' Fractional bin hours for different bin temperatures for region three, from ANSI/AHRI 210/240

RegionThreeFracBinHoursAtOutdoorBinTemp = Array(0.153, 0.142, 0.138, 0.137, 0.135, 0.118, _
0.092, 0.047, 0.021, 0.009, 0.005, 0.002, 0.001, 0#, 0#, 0#, 0#, 0#)
' Fractional bin hours for different bin temperatures for region four, from ANSI/AHRI 210/240

RegionFourFracBinHoursAtOutdoorBinTemp = Array(0.132, 0.111, 0.103, 0.093, 0.1, 0.109, _
0.126, 0.087, 0.055, 0.036, 0.026, 0.013, 0.006, 0.002, 0.001, 0#, 0#, 0#)
' Fractional bin hours for different bin temperatures for region five, from ANSI/AHRI 210/240

RegionFiveFracBinHoursAtOutdoorBinTemp = Array(0.106, 0.092, 0.086, 0.076, 0.078, 0.087, _
0.102, 0.094, 0.074, 0.055, 0.047, 0.038, 0.029, 0.018, 0.01, 0.005, 0.002, 0.001)
' Fractional bin hours for different bin temperatures for region six, from ANSI/AHRI 210/240

RegionSixFracBinHoursAtOutdoorBinTemp = Array(0.113, 0.206, 0.215, 0.204, 0.141, 0.076, _
0.034, 0.008, 0.003, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)

End Sub

Function MainHSPF(RegionNumber As Integer, CapFTCurveType As String, EIRFTCurveType As String, CapFFCurveType As String, EIRFFCurveType As String, DefrostControlOption As String, CurrentRowNum As Integer)

' PURPOSE OF THIS SUBROUTINE:
' Find the value of x between x0 and x1 such that HSPFResidual(X)
' is equal to zero.

' METHODOLOGY EMPLOYED:
' Uses the Regula Falsi (false position) method (similar to secant method)

' REFERENCES:
' See Press et al., Numerical Recipes in Fortran, Cambridge University Press,
' 2nd edition, 1992. Page 347 ff.

' SUBROUTINE PARAMETER DEFINITIONS:
Dim Eps As Single ' required absolute accuracy
Dim MaxIte As Integer ' maximum number of allowed iterations
Dim Flag As Integer ' integer storing exit status
' = -2: HSPFResidual(X0) and HSPFResidual(X1) have the same sign
' = -1: no convergence
' >  0: number of iterations performed
Dim XRes As Single ' value of x that solves HSPFResidual(X [,Par]) = 0
Dim X_0 As Single ' 1st bound of interval that contains the solution
Dim X_1 As Single ' 2nd bound of interval that contains the solution
Dim Small As Double
Small = 0.0000000001
Eps = 0.0001 ' required absolute accuracy
MaxIte = 500 ' maximum number of allowed iterations

' DERIVED TYPE DEFINITIONS
' na

' SUBROUTINE LOCAL VARIABLE DECLARATIONS:
Dim X0 As Single ' present 1st bound
Dim X1 As Single ' present 2nd bound
Dim XTemp As Single ' new estimate
Dim Y0 As Single ' f at X0
Dim Y1 As Single ' f at X1
Dim YTemp As Single ' f at XTemp
Dim DY As Single ' DY = Y0 - Y1
Dim Conv As Boolean ' flag, true if convergence is achieved
Dim StopMaxIte As Boolean ' stop due to exceeding of maximum # of iterations
Dim Cont As Boolean ' flag, if true, continue searching
Dim NIte As Integer ' number of interations

X_0 = 1      ' Minimum COP value for iterative solutions
X_1 = 10    ' Maximum COP value for iterative solutions

RegionNum = RegionNumber
CapFTempCurveType = CapFTCurveType
EIRFTempCurveType = EIRFTCurveType
CapFFlowCurveType = CapFFCurveType
EIRFFlowCurveType = EIRFFCurveType
DefrostControl = DefrostControlOption
RowNum = CurrentRowNum

X0 = X_0
X1 = X_1
Conv = False
StopMaxIte = False
Cont = True
NIte = 0

If CheckInputs = True Then
    Call SetGlobalConstantArrays
    Y0 = HSPFResidual(X0)
    Y1 = HSPFResidual(X1)
    ' check initial values
     If (Y0 * Y1 > 0) Then
       Flag = -2
       XRes = X0
       'Return
       Cont = False
    End If
Else
  Cont = False
End If

If LastHP = False Then
    MainHSPF = True
Else
    MainHSPF = False
End If
 
Do While (Cont = True)

  DY = Y0 - Y1
  If (Abs(DY) < Small) Then DY = Small
    ' new estimation
    XTemp = (Y0 * X1 - Y1 * X0) / DY
    YTemp = HSPFResidual(XTemp)

    NIte = NIte + 1

    ' check convergence
    If (Abs(YTemp) < Eps) Then Conv = True

      If (NIte > MaxIte) Then StopMaxIte = True

        If ((Conv = False) And (StopMaxIte = False)) Then
          Cont = True
        Else
          Cont = False
        End If

        If (Cont) Then

          ' reassign values (only if further iteration required)
          If (Y0 < 0#) Then
            If (YTemp < 0#) Then
              X0 = XTemp
              Y0 = YTemp
            Else
              X1 = XTemp
              Y1 = YTemp
            End If
          Else
            If (YTemp < 0#) Then
              X1 = XTemp
              Y1 = YTemp
            Else
              X0 = XTemp
              Y0 = YTemp
            End If
         End If ' ( Y0 < 0 )

       End If ' (Cont)

  Loop ' Cont

  If (Conv) Then
    Flag = NIte
  Else
    Flag = -1
  End If
  XRes = XTemp
  
End Function
Function CalcHSPF(RatedCOP As Single)
' PURPOSE OF THIS SUB:
'     Calculates Heating Seasonal Performance Factor (HSPF) for Air-Source Direct Expansion Heat Pumps having
'     a single-speed compressor, fixed speed indoor supply air fan

' METHODOLOGY EMPLOYED:
' Methodology for calculating standard ratings for DX air air source heat pumps
'     (1) Obtains the rated condition parameters:
'         Heating capacity, COP,  Rated Air volume flow rate through the
'         DX Cooling Coil and Fan power per rated air flow rate
'
'     (2) Evaluates the heating coil capacities for AHRI tests H1, H2 and H3 using the performance cuves and
'         input values specified at (1) above. Then net heating capacity is determined from the total heating capacity
'         of the DX coil at the AHRI test conditions and accounting for the INDOOR supply air fan heat.
'
'     (3) Calculates the electric power consumed by the DX Coil Unit (compressor + outdoor condenser fan).
'         The net electric power consumption is determined by adding the indoor fan electric power to the
'         electric power consumption by the DX Coil Condenser Fan and Compressor at the AHRI test conditions.
'
'     (4) High Temperature Heating Standard (Net) Rating Capacity and Low Temperature Heating Standard (Net)
'         Rating Capacity capacity are determined using tests H1 adn H3 per ANSI/AHRI 210/240 2008.
'
'     (5) The HSPF is evaluated from the total net heating capacity and total electric power
'         evaluated at the standard rated test conditions. For user specified region number, the outdoor temperatures
'         are Binned (grouped) and fractioanl bin hours for each bin over the entire heating season are taken
'         from AHRI 210/240. Then for each bin, building load, heat pump energy adn resistance space heating enegry are
'         calculated. The sum of building load divided by sum of heat pump and resistance space heating over the
'         entire heating season gives the HSPF. The detailed calculation algorithms of calculating HSPF
'         are described in Engineering Reference.
'
' REFERENCES:
' (1) ANSI/AHRI Standard 210/240-2008:  Standard for Performance Rating of Unitary Air-Conditioning and
'           Air-Source Heat Pumps. Arlington, VA:  Air-Conditioning, Heating
'           , and Refrigeration Institute.

' SUB LOCAL VARIABLE DECLARATIONS:
Dim TotalHeatingCapRated As Double '= 0#      ' Heating Coil capacity at Rated conditions, without accounting supply fan heat [W]
Dim EIRRated As Double ' = 0#  ' EIR at Rated conditions [-]
Dim TotCapTempModFacRated As Double ' = 0#     ' Total capacity as a function of temerature modifier at rated conditions [-]
Dim EIRTempModFacRated As Double ' = 0#        ' EIR as a function of temerature modifier at rated conditions [-]
Dim TotalHeatingCapH2Test As Double ' = 0#     ' Heating Coil capacity at H2 test conditions, without accounting supply fan heat [W]
Dim TotalHeatingCapH3Test As Double ' = 0#     ' Heating Coil capacity at H3 test conditions, without accounting supply fan heat [W]
Dim NetHeatingCapRated As Double ' = 0#     ' Net heating Coil capacity at Rated conditions W]
Dim CapTempModFacH2Test As Double '= 0#       ' Total capacity as a function of temerature modifier at H2 test conditions [-]
Dim EIRTempModFacH2Test As Double ' = 0#       ' EIR as a function of temerature modifier at H2 test conditions [-]
Dim EIRH2Test As Double ' = 0# ' EIR at H2 test conditions [-]
Dim CapTempModFacH3Test As Double '= 0#       ' Total capacity as a function of temerature modifier at H3 test conditions [-]
Dim EIRTempModFacH3Test As Double '= 0#       ' EIR as a function of temerature modifier at H3 test conditions [-]
Dim EIRH3Test As Double '= 0# ' EIR at H3 test conditions [-]
Dim TotCapFlowModFac As Double ' = 0#          ' Total capacity modifier(function of actual supply air flow vs rated flow)
Dim EIRFlowModFac As Double ' = 0#               ' EIR modifier(function of actual supply air flow vs rated flow)
Dim FanPowerPerEvapAirFlowRate As Double ' = 0#    ' Fan power per air volume flow rate [W/(m3/s)]

Dim ElecPowerRated As Double '' Total system power at Rated conditions accounting for supply fan heat [W]
Dim ElecPowerH2Test As Double '                  ' Total system power at H2 test conditions accounting for supply fan heat [W]
Dim ElecPowerH3Test As Double '                   ' Total system power at H3 test conditions accounting for supply fan heat [W]
Dim NetHeatingCapH2Test As Double '           ' Net Heating Coil capacity at H2 test conditions accounting for supply fan heat [W]
Dim NetHeatingCapH3Test As Double '           ' Net Heating Coil capacity at H3 test conditions accounting for supply fan heat [W]

Dim PartLoadFactor As Double '
Dim LoadFactor As Double '     ' Frac. "on" time for last stage at the desired reduced capacity, (dimensionless)
Dim LowTempCutOutFactor As Double ' = 0#       ' Factor which corresponds to compressor operation depending on outdoor temperature
Dim OATempCompressorOff As Double '= 0#       ' Minimum outdoor air temperature to turn the commpressor off, [C]
Dim OATempCompressorOn As Double '= 0#       ' The outdoor tempearture when the compressor is automatically turned
'back on, if applicable, following automatic shut off. This field is
'used only for HSPF calculation. [C]
Dim FractionalBinHours As Double ' = 0#        ' Fractional bin hours for the heating season [-]
Dim BuildingLoad As Double ' = 0#                  ' Building space conditioning load corresponding to an outdoor bin temperature [W]
Dim HeatingModeLoadFactor As Double '= 0#     ' Heating mode load factor corresponding to an outdoor bin temperature [-]
Dim NetHeatingCapReduced As Double ' = 0#      ' Net Heating Coil capacity corresponding to an outdoor bin temperature [W]
Dim TotalBuildingLoad As Double '= 0#         ' Sum of building load over the entire heating season [W]
Dim TotalElectricalEnergy As Double '= 0#     ' Sum of electrical energy consumed by the heatpump over the heating season [W]
Dim DemandDeforstCredit As Double '= 1#       ' A factor to adjust HSPF if coil has demand defrost control [-]
Dim CheckCOP As Double '= 0#  ' Checking COP at an outdoor bin temperature against unity [-]
Dim DesignHeatingRequirement As Double '= 0#         ' The amount of heating required to maintain a given indoor temperature
' at a particular outdoor design temperature.  [W]
Dim DesignHeatingRequirementMin As Double ' = 0#      ' minimum design heating requirement [W]
Dim DesignHeatingRequirementMax As Double '= 0#      ' maximum design heating requirement [W]
Dim ElectricalPowerConsumption As Double '= 0#       ' Electrical power corresponding to an outdoor bin temperature [W]
Dim HeatPumpElectricalEnergy As Double ' = 0#         ' Heatpump electrical energy corresponding to an outdoor bin temperature [W]
Dim TotalHeatPumpElectricalEnergy As Double '= 0#    ' Sum of Heatpump electrical energy over the entire heating season [W]
Dim ResistiveSpaceHeatingElectricalEnergy As Double '= 0#         ' resistance heating electrical energy corresponding to an
' outdoor bin temperature [W]
Dim TotalResistiveSpaceHeatingElectricalEnergy As Double '= 0#    ' Sum of resistance heating electrical energy over the
' entire heating season [W]

Dim BinNum As Integer ' bin number counter
Dim StandardDHRNum As Integer ' Integer counter for standardized DHRs
Dim FanPowerPerEvapAirFlowRateFromInput As Single
Dim RatedTotalCapacity As Single
Dim RatedAirVolFlowRate As Single
Dim MinOATCompressor As Single
Dim OATempCompressorOnOffBlank As Boolean ' Flag used to determine low temperature cut out factor

If (IsEmpty(ThisWorkbook.Sheets("HSPF").Cells(RowNum, OATempCompressorOn_ColumnNum).Value) = True) Then
  OATempCompressorOnOffBlank = True
Else
  OATempCompressorOnOffBlank = False
End If

TotalBuildingLoad = 0
TotalHeatPumpElectricalEnergy = 0
TotalResistiveSpaceHeatingElectricalEnergy = 0

' Calculate the supply air fan electric power consumption.  The electric power consumption is estimated
' using either user supplied or AHRI default value for fan power per air volume flow rate
FanPowerPerEvapAirFlowRateFromInput = ThisWorkbook.Sheets("HSPF").Cells(RowNum, EvapFanPowerPerVolFlowRate_ColumnNum).Value
RatedTotalCapacity = ThisWorkbook.Sheets("HSPF").Cells(RowNum, RatedTotalHeatingCapacity_ColumnNum).Value
RatedAirVolFlowRate = ThisWorkbook.Sheets("HSPF").Cells(RowNum, RatedAirFlowRate_ColumnNum).Value
MinOATCompressor = ThisWorkbook.Sheets("HSPF").Cells(RowNum, MinOATCompressor_ColumnNum).Value
OATempCompressorOn = ThisWorkbook.Sheets("HSPF").Cells(RowNum, OATempCompressorOn_ColumnNum).Value

If (FanPowerPerEvapAirFlowRateFromInput <= 0#) Then
  FanPowerPerEvapAirFlowRate = DefaultFanPowerPerEvapAirFlowRate
Else
  FanPowerPerEvapAirFlowRate = FanPowerPerEvapAirFlowRateFromInput
End If

TotCapFlowModFac = CurveValue(CapFFlowCurveType, CapFFF, RowNum, AirMassFlowRatioRated)
EIRFlowModFac = CurveValue(EIRFFlowCurveType, EIRFFF, RowNum, AirMassFlowRatioRated)

Select Case (CapFTempCurveType)

  Case ("Quadratic")
    TotCapTempModFacRated = CurveValue(CapFTempCurveType, CapFT, RowNum, HeatingOutdoorCoilInletAirDBTempRated)

    CapTempModFacH2Test = CurveValue(CapFTempCurveType, CapFT, RowNum, HeatingOutdoorCoilInletAirDBTempH2Test)

    CapTempModFacH3Test = CurveValue(CapFTempCurveType, CapFT, RowNum, HeatingOutdoorCoilInletAirDBTempH3Test)
  Case ("Cubic")
    TotCapTempModFacRated = CurveValue(CapFTempCurveType, CapFT, RowNum, HeatingOutdoorCoilInletAirDBTempRated)

    CapTempModFacH2Test = CurveValue(CapFTempCurveType, CapFT, RowNum, HeatingOutdoorCoilInletAirDBTempH2Test)

    CapTempModFacH3Test = CurveValue(CapFTempCurveType, CapFT, RowNum, HeatingOutdoorCoilInletAirDBTempH3Test)
  Case ("Biquadratic")
    TotCapTempModFacRated = CurveValue(CapFTempCurveType, CapFT, RowNum, HeatingIndoorCoilInletAirDBTempRated, _
    HeatingOutdoorCoilInletAirDBTempRated)

    CapTempModFacH2Test = CurveValue(CapFTempCurveType, CapFT, RowNum, HeatingIndoorCoilInletAirDBTempRated, _
    HeatingOutdoorCoilInletAirDBTempH2Test)

    CapTempModFacH3Test = CurveValue(CapFTempCurveType, CapFT, RowNum, HeatingIndoorCoilInletAirDBTempRated, _
    HeatingOutdoorCoilInletAirDBTempH3Test)

End Select

Select Case (EIRFTempCurveType)

  Case ("Quadratic")
    EIRTempModFacRated = CurveValue(EIRFTempCurveType, EIRFT, RowNum, HeatingOutdoorCoilInletAirDBTempRated)

    EIRTempModFacH2Test = CurveValue(EIRFTempCurveType, EIRFT, RowNum, HeatingOutdoorCoilInletAirDBTempH2Test)

    EIRTempModFacH3Test = CurveValue(EIRFTempCurveType, EIRFT, RowNum, HeatingOutdoorCoilInletAirDBTempH3Test)
  Case ("Cubic")
    EIRTempModFacRated = CurveValue(EIRFTempCurveType, EIRFT, RowNum, HeatingOutdoorCoilInletAirDBTempRated)

    EIRTempModFacH2Test = CurveValue(EIRFTempCurveType, EIRFT, RowNum, HeatingOutdoorCoilInletAirDBTempH2Test)

    EIRTempModFacH3Test = CurveValue(EIRFTempCurveType, EIRFT, RowNum, HeatingOutdoorCoilInletAirDBTempH3Test)
  Case ("Biquadratic")
    EIRTempModFacRated = CurveValue(EIRFTempCurveType, EIRFT, RowNum, HeatingIndoorCoilInletAirDBTempRated, _
    HeatingOutdoorCoilInletAirDBTempRated)

    EIRTempModFacH2Test = CurveValue(EIRFTempCurveType, EIRFT, RowNum, HeatingIndoorCoilInletAirDBTempRated, _
    HeatingOutdoorCoilInletAirDBTempH2Test)

    EIRTempModFacH3Test = CurveValue(EIRFTempCurveType, EIRFT, RowNum, HeatingIndoorCoilInletAirDBTempRated, _
    HeatingOutdoorCoilInletAirDBTempH3Test)
End Select

TotalHeatingCapRated = RatedTotalCapacity * TotCapTempModFacRated * TotCapFlowModFac
NetHeatingCapRated = TotalHeatingCapRated + FanPowerPerEvapAirFlowRate * RatedAirVolFlowRate

TotalHeatingCapH2Test = RatedTotalCapacity * CapTempModFacH2Test * TotCapFlowModFac
NetHeatingCapH2Test = TotalHeatingCapH2Test + FanPowerPerEvapAirFlowRate * RatedAirVolFlowRate

TotalHeatingCapH3Test = RatedTotalCapacity * CapTempModFacH3Test * TotCapFlowModFac
NetHeatingCapH3Test = TotalHeatingCapH3Test + FanPowerPerEvapAirFlowRate * RatedAirVolFlowRate

If (RegionNum = 5) Then
  DesignHeatingRequirementMin = NetHeatingCapRated
Else
  DesignHeatingRequirementMin = NetHeatingCapRated * 1.8 * (18.33 - OutdoorDesignTemperature(RegionNum - 1)) / (60)
End If

For StandardDHRNum = 0 To TotalNumOfStandardDHRs - 2
  If (StandardDesignHeatingRequirement(StandardDHRNum) <= DesignHeatingRequirementMin And _
    StandardDesignHeatingRequirement(StandardDHRNum + 1) >= DesignHeatingRequirementMin) Then
    If ((DesignHeatingRequirementMin - StandardDesignHeatingRequirement(StandardDHRNum)) > _
      (StandardDesignHeatingRequirement(StandardDHRNum + 1) - DesignHeatingRequirementMin)) Then
      DesignHeatingRequirementMin = StandardDesignHeatingRequirement(StandardDHRNum + 1)
    Else
      DesignHeatingRequirementMin = StandardDesignHeatingRequirement(StandardDHRNum)
    End If
  End If
Next

If (StandardDesignHeatingRequirement(0) >= DesignHeatingRequirementMin) Then
  DesignHeatingRequirement = StandardDesignHeatingRequirement(0)
ElseIf (StandardDesignHeatingRequirement(TotalNumOfStandardDHRs - 1) <= DesignHeatingRequirementMin) Then
  DesignHeatingRequirement = StandardDesignHeatingRequirement(TotalNumOfStandardDHRs - 1)
Else
  DesignHeatingRequirement = DesignHeatingRequirementMin
End If

For BinNum = 0 To TotalNumOfTemperatureBins(RegionNum - 1) - 1

If (RegionNum = 1) Then
  FractionalBinHours = RegionOneFracBinHoursAtOutdoorBinTemp(BinNum)
ElseIf (RegionNum = 2) Then
  FractionalBinHours = RegionTwoFracBinHoursAtOutdoorBinTemp(BinNum)
ElseIf (RegionNum = 3) Then
  FractionalBinHours = RegionThreeFracBinHoursAtOutdoorBinTemp(BinNum)
ElseIf (RegionNum = 4) Then
  FractionalBinHours = RegionFourFracBinHoursAtOutdoorBinTemp(BinNum)
ElseIf (RegionNum = 5) Then
  FractionalBinHours = RegionFiveFracBinHoursAtOutdoorBinTemp(BinNum)
ElseIf (RegionNum = 6) Then
  FractionalBinHours = RegionSixFracBinHoursAtOutdoorBinTemp(BinNum)
End If

BuildingLoad = (18.33 - OutdoorBinTemperature(BinNum)) / (18.33 - OutdoorDesignTemperature(RegionNum - 1)) _
* CorrectionFactor * DesignHeatingRequirement

If (DefrostControl = "Timed") Then
  DemandDeforstCredit = 1# ' Timed defrost control
Else
  DemandDeforstCredit = 1.03 ' Demand defrost control
End If

OATempCompressorOff = MinOATCompressor

If (RatedCOP > 0#) Then      ' RatedCOP <= 0.0 is trapped in GetInput, but keep this as "safety"

  EIRRated = EIRTempModFacRated * EIRFlowModFac / RatedCOP
  EIRH2Test = EIRTempModFacH2Test * EIRFlowModFac / RatedCOP
  EIRH3Test = EIRTempModFacH3Test * EIRFlowModFac / RatedCOP

End If

ElecPowerRated = EIRRated * TotalHeatingCapRated + FanPowerPerEvapAirFlowRate * RatedAirVolFlowRate
ElecPowerH2Test = EIRH2Test * TotalHeatingCapH2Test + FanPowerPerEvapAirFlowRate * RatedAirVolFlowRate
ElecPowerH3Test = EIRH3Test * TotalHeatingCapH3Test + FanPowerPerEvapAirFlowRate * RatedAirVolFlowRate

If ((OutdoorBinTemperature(BinNum) <= -8.33) Or (OutdoorBinTemperature(BinNum) >= 7.22)) Then
  NetHeatingCapReduced = NetHeatingCapH3Test + (NetHeatingCapRated - NetHeatingCapH3Test) * _
  (OutdoorBinTemperature(BinNum) + 8.33) / (16.67)
  ElectricalPowerConsumption = ElecPowerH3Test + (ElecPowerRated - ElecPowerH3Test) * _
  (OutdoorBinTemperature(BinNum) + 8.33) / (16.67)
Else
  NetHeatingCapReduced = NetHeatingCapH3Test + (NetHeatingCapH2Test - NetHeatingCapH3Test) * _
  (OutdoorBinTemperature(BinNum) + 8.33) / (10)
  ElectricalPowerConsumption = ElecPowerH3Test + (ElecPowerH2Test - ElecPowerH3Test) * _
  (OutdoorBinTemperature(BinNum) + 8.33) / (10)
End If

If (NetHeatingCapReduced <> 0) Then
  HeatingModeLoadFactor = BuildingLoad / NetHeatingCapReduced
End If

If (HeatingModeLoadFactor > 1) Then
  HeatingModeLoadFactor = 1
End If

PartLoadFactor = 1 - CyclicDegradationCoeff * (1 - HeatingModeLoadFactor)

If (ElectricalPowerConsumption <> 0) Then
  CheckCOP = NetHeatingCapReduced / ElectricalPowerConsumption
End If

If (CheckCOP < 1#) Then
  LowTempCutOutFactor = 0
Else
  If (OATempCompressorOnOffBlank = False) Then
    If (OutdoorBinTemperature(BinNum) <= OATempCompressorOff) Then
      LowTempCutOutFactor = 0
    ElseIf (OutdoorBinTemperature(BinNum) > OATempCompressorOff And _
      OutdoorBinTemperature(BinNum) <= OATempCompressorOn) Then
      LowTempCutOutFactor = 0.5
    Else
      LowTempCutOutFactor = 1
    End If
  Else
    LowTempCutOutFactor = 1
  End If
End If

If (PartLoadFactor <> 0) Then
  HeatPumpElectricalEnergy = (HeatingModeLoadFactor * ElectricalPowerConsumption * LowTempCutOutFactor) _
  * FractionalBinHours / PartLoadFactor
End If

ResistiveSpaceHeatingElectricalEnergy = (BuildingLoad - HeatingModeLoadFactor * NetHeatingCapReduced _
* LowTempCutOutFactor) * FractionalBinHours

TotalBuildingLoad = TotalBuildingLoad + (BuildingLoad * FractionalBinHours)

TotalHeatPumpElectricalEnergy = TotalHeatPumpElectricalEnergy + HeatPumpElectricalEnergy

TotalResistiveSpaceHeatingElectricalEnergy = TotalResistiveSpaceHeatingElectricalEnergy + _
ResistiveSpaceHeatingElectricalEnergy
Next

TotalElectricalEnergy = TotalHeatPumpElectricalEnergy + TotalResistiveSpaceHeatingElectricalEnergy

If (TotalElectricalEnergy <> 0#) Then
  CalcHSPF = TotalBuildingLoad * DemandDeforstCredit / TotalElectricalEnergy * ConvFromSIToIP
End If

ThisWorkbook.Sheets("HSPF").Cells(RowNum, CalculatedHSPF_ColumnNum).Value = CalcHSPF
ThisWorkbook.Sheets("HSPF").Cells(RowNum, COP_ColumnNum).Value = RatedCOP

'Return
End Function

Function CurveValue(CurveType As String, CurveNameIndex As Integer, RowNum As Integer, Arg1 As Double, Optional Arg2 As Double) As Double

  Dim C1 As Double
  Dim C2 As Double
  Dim C3 As Double
  Dim C4 As Double
  Dim C5 As Double
  Dim C6 As Double

  Dim FirstColumNumForCoeff As Integer

  If CurveNameIndex = 1 Then '
    FirstColumNumForCoeff = CapFT_C1_ColumnNum
  ElseIf CurveNameIndex = 2 Then
    FirstColumNumForCoeff = EIRFT_C1_ColumnNum
  ElseIf CurveNameIndex = 3 Then
    FirstColumNumForCoeff = CapFFF_C1_ColumnNum
  ElseIf CurveNameIndex = 4 Then
    FirstColumNumForCoeff = EIRFFF_C1_ColumnNum
  Else
    ' Should never come here
  End If

  C1 = ThisWorkbook.Sheets("HSPF").Cells(RowNum, FirstColumNumForCoeff).Value
  C2 = ThisWorkbook.Sheets("HSPF").Cells(RowNum, FirstColumNumForCoeff + 1).Value
  C3 = ThisWorkbook.Sheets("HSPF").Cells(RowNum, FirstColumNumForCoeff + 2).Value
  C4 = ThisWorkbook.Sheets("HSPF").Cells(RowNum, FirstColumNumForCoeff + 3).Value
  C5 = ThisWorkbook.Sheets("HSPF").Cells(RowNum, FirstColumNumForCoeff + 4).Value
  C6 = ThisWorkbook.Sheets("HSPF").Cells(RowNum, FirstColumNumForCoeff + 5).Value

  If CurveType = "Quadratic" Then
    CurveValue = C1 + C2 * Arg1 + C3 * (Arg1 ^ 2)
  ElseIf CurveType = "Cubic" Then
    CurveValue = C1 + C2 * Arg1 + C3 * (Arg1 ^ 2) + C4 * (Arg1 ^ 3)
  ElseIf CurveType = "Biquadratic" Then
    CurveValue = C1 + C2 * Arg1 + C3 * (Arg1 ^ 2) + C4 * Arg2 + C5 * (Arg2 ^ 2) + C6 * Arg1 * Arg2
  Else
    ' Should never come here
  End If

End Function

Function HSPFResidual(COP As Single)

          ' PURPOSE OF THIS FUNCTION:
          ' Calculates residual function (Desired HSPF - Calculated HSPF) / Desired HSPF.
          ' Calculated HSPF depends on the COP which is being varied to zero the residual.

          ' METHODOLOGY EMPLOYED:
          ' Calls CalcHSPF with COP specified here, and calculates
          ' the residual as defined above.

          ' SUBROUTINE ARGUMENT DEFINITIONS:
  Dim DesiredHSPF As Single          ' residual to be minimized to zero
  Dim CurrentRowNum As Integer

  DesiredHSPF = ThisWorkbook.Sheets("HSPF").Cells(RowNum, DesiredHSPF_ColumnNum).Value
  HSPFResidual = (DesiredHSPF - CalcHSPF(COP)) / DesiredHSPF

End Function

Function CheckInputs()

CheckInputs = True
LastHP = False

If (IsEmpty(ThisWorkbook.Sheets("HSPF").Cells(RowNum, CoolingCapacity_ColumnNum).Value) = True And _
    IsEmpty(ThisWorkbook.Sheets("HSPF").Cells(RowNum, RatedTotalHeatingCapacity_ColumnNum).Value) = True And _
    IsEmpty(ThisWorkbook.Sheets("HSPF").Cells(RowNum, RatedAirFlowRate_ColumnNum).Value) = True) Then
    LastHP = True
    CheckInputs = False
End If

If (CheckInputs = False And LastHP = False) Then
  ThisWorkbook.Sheets("HSPF").Cells(RowNum, DesiredHSPF_ColumnNum).Value = "NA"
  ThisWorkbook.Sheets("HSPF").Cells(RowNum, CalculatedHSPF_ColumnNum).Value = "--"
  ThisWorkbook.Sheets("HSPF").Cells(RowNum, COP_ColumnNum).Value = "--"
  MsgBox "Check inputs. Program is exiting."
  End
End If
End Function

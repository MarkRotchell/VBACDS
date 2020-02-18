Attribute VB_Name = "CDS"
Option Explicit

'#################################################################################################################
'# Module for Calculating CDS valuation according to standard ISDA model and standard ISDA contract specification
'#
'# Written by Mark Rotchell
'# https://rotchvba.wordpress.com/
'#
'# Key Worksheet Functions
'#  - MR_CDS_Valuation: Calculates CDS NPV when hazard rate is known
'#  - MR_CDS_NPVFromTradedSpread: Calculates CDS NPV when traded spread is known
'#
'# Key Assumptions
'#  - premiums are quarterly act/360
'#  - protection is start of day
'#  - accrued interest is paid on default
'#  - premium is accrued for maturity date
'#  - maturity date is unadjusted
'#
'#
'# Input Requirements
'#  - Coupon, Recovery, Interest Rates, Hazard rates and traded spread should be entered as absolute numbers, not bps,
'#     i.e., a coupon of 100bp should be entered as 0.01
'#  - Yield and hazard rates are continuously compounded zero-coupon rates act/365
'#  - Yield curve must have sorted and unique dates
'#
'# Coding Philosophy
'#  - This code is designed to be accessible to the casual VBA user interested in CDS
'#  - Most VBA users are not professional programmers, so keep things as simple as possible, in particular:
'#    - No classes
'#    - No windows API calls
'#
'#  - Many VBA users do not have permissions to install addins, or even download certain types of files
'#    - code should be "Copy-And-Pastable" as text - no need to download or install addins or files containing macros
'#    - no references to external libraries needed - everying in "pure" VBA and Excel
'#    - everything in one module: end user can split if they like, but easier to copy-and-paste if not
'#
'# TODO
'#  - The bootstrapping of the Yield Curve is the next task. For now this is just focused on the credit peice.
'#
'#################################################################################################################


Private Const SPOT_DAYS  As Long = 2
Private Const CASH_DAYS  As Long = 3
Private Const NO_ERROR As Long = 0
Private Const QUARTERS_IN_YEAR As Long = 4

'########################################################################################
'# Data Structures
'########################################################################################

' A Yield Curve
Private Type YieldMarketData
    CurveDate As Date
    Dates As Collection
    Rates As Collection
End Type

' Credit market data
Private Type CreditMarketData
    CurveDate As Date
    HazardRate As Double
    TradedSpread As Double
    Recovery As Double
End Type

' Dates needed to value a CDS
Private Type CDSValuationDates
    PaymentDates As Collection
    AccrualDates As Collection
    ProtectionNodes As Collection
    AccruedOnDefaultNodes As Collection
    AccruedOnDefaultAccrualStartDates As Collection
End Type

' Information about a CDS contract
Private Type CDSContractSpecification
    TradeDate As Date
    CashDate As Date
    MaturityDate As Date
    Notional As Double
    Coupon As Double
End Type

' Some node or interval in a CDS valuation
Private Type ValuationNode
    AccrualStart As Date
    ProtectionDate As Date
    AccrualDate As Date
    PaymentDate As Date
    Zero As Double
    DiscountFactor As Double
    SurvivalProbability As Double
    LogDF As Double
    RiskyDF As Double
    Interest As Double
    Hazard As Double
    RiskyDFChange As Double
    StartDCC As Double
    PeriodDCC As Double
    PV As Double
End Type

'########################################################################################
'# WorkSheet Functions
'########################################################################################

Public Function MR_CDS_Valuation( _
        ByVal TradeDate As Date, _
        ByVal MaturityDate As Date, _
        ByVal Notional As Double, _
        ByVal Coupon As Double, _
        ByVal Recovery As Double, _
        ByVal HazardRate As Double, _
        ByRef YieldDates As Range, _
        ByRef YieldRates As Range, _
        Optional ByRef isDirty As Boolean = True) _
        As Variant
    
    ' WORKSHEET FUNCTION
    ' Calculates the NPV of a CDS contract when the Hazard Rate is known
    ' Coupon, Recovery and Hazard rate should be entered as absolute numbers, not bps. i.e. a coupon of 100bp should be entered as 0.01
    ' Yield Rates should be Continuously Compounded ACT/365
    
    Dim InputDataQuality As XlCVError
    Dim SpotDate         As Date
    Dim YieldMarket      As YieldMarketData
    Dim CreditMarket     As CreditMarketData
    Dim CDSContract      As CDSContractSpecification
    Dim ValuationDates   As CDSValuationDates
    
    InputDataQuality = ValidateInputData(TradeDate, MaturityDate, Recovery, HazardRate, Coupon, 0, YieldDates, YieldRates)
    
    If InputDataQuality <> NO_ERROR Then
        MR_CDS_Valuation = InputDataQuality
    Else
        SpotDate = Application.WorkDay(TradeDate, SPOT_DAYS)
        YieldMarket = LoadYieldMarket(SpotDate, YieldDates, YieldRates)
        CreditMarket = LoadCreditMarket(TradeDate, Recovery, HazardRate)
        CDSContract = LoadCDSContract(TradeDate, MaturityDate, Notional, Coupon)
        ValuationDates = GenerateValuationDates(CDSContract, YieldMarket)
        MR_CDS_Valuation = CDSNPV(CDSContract, ValuationDates, YieldMarket, CreditMarket, isDirty)
    End If
End Function

Public Function MR_CDS_AccruedInterest( _
        ByVal TradeDate As Date, _
        ByVal MaturityDate As Date, _
        ByVal Notional As Double, _
        ByVal Coupon As Double) _
        As Variant
    
    ' WORKSHEET FUNCTION
    ' Calculates the accrued interest of a CDS contract
    ' Coupon should be entered as absolute number, not bps. i.e. a coupon of 100bp should be entered as 0.01
    
    Application.Volatile False
    
    Dim Unadjusted As Date
    Dim Adjusted   As Date
    Dim AccrualStart As Date
    
    Unadjusted = MaturityDate
    Adjusted = Unadjusted
    
    Do While Adjusted > TradeDate
        Unadjusted = Application.WorksheetFunction.CoupPcd(Unadjusted - 1, MaturityDate, QUARTERS_IN_YEAR)
        Adjusted = ModFol(Unadjusted)
    Loop
    
    AccrualStart = Adjusted - 1
    MR_CDS_AccruedInterest = Notional * Coupon * Act360(AccrualStart, TradeDate)
    
End Function


Public Function MR_CDS_NPVFromTradedSpread( _
        ByVal TradeDate As Date, _
        ByVal MaturityDate As Date, _
        ByVal Notional As Double, _
        ByVal Coupon As Double, _
        ByVal Recovery As Double, _
        ByVal TradedSpread As Double, _
        ByRef YieldDates As Range, _
        ByRef YieldRates As Range, _
        Optional ByRef isDirty As Boolean = True) _
        As Variant
        
    ' WORKSHEET FUNCTION
    ' Calculates the NPV of a CDS using the Traded Spread
    ' Coupon, Recovery and Traded Spread should be entered as absolute numbers, not bps. i.e. a coupon of 100bp should be entered as 0.01
    ' Yield Rates should be Continuously Compounded ACT/365
    
    Application.Volatile False
    
    Dim InputDataQuality As XlCVError
    Dim SpotDate         As Date
    Dim YieldMarket      As YieldMarketData
    Dim CreditMarket     As CreditMarketData
    Dim CDSContract      As CDSContractSpecification
    Dim ValuationDates   As CDSValuationDates
    
    InputDataQuality = ValidateInputData(TradeDate, MaturityDate, Recovery, 0, Coupon, TradedSpread, YieldDates, YieldRates)
    
    If InputDataQuality <> NO_ERROR Then
        MR_CDS_NPVFromTradedSpread = InputDataQuality
    Else
        SpotDate = Application.WorkDay(TradeDate, SPOT_DAYS)
        YieldMarket = LoadYieldMarket(SpotDate, YieldDates, YieldRates)
        CreditMarket = LoadCreditMarket(TradeDate, Recovery, , TradedSpread)
        CDSContract = LoadCDSContract(TradeDate, MaturityDate, Notional, Coupon)
        ValuationDates = GenerateValuationDates(CDSContract, YieldMarket)
        MR_CDS_NPVFromTradedSpread = CDSNPVfromTradedSpread(CDSContract, ValuationDates, YieldMarket, CreditMarket, isDirty)
    End If
    
End Function



'########################################################################################
'# Data Validation for Inputs from Worksheet Functions
'########################################################################################

Private Function ValidateInputData( _
        ByVal TradeDate As Date, _
        ByVal MaturityDate As Date, _
        ByVal Recovery As Double, _
        ByVal HazardRate As Double, _
        ByVal Coupon As Double, _
        ByVal TradedSpread As Double, _
        ByRef YieldDatesRange As Range, _
        ByRef YieldRatesRange As Range) _
    As XlCVError

    ' validate data input by user
    
    Dim ErrorType         As XlCVError
    Dim YieldDatesArray() As Variant
    Dim YieldRatesArray() As Variant
    Dim YieldDates        As New Collection
    Dim YieldRates        As New Collection
    Dim thisDate          As Variant
    Dim lastDate          As Variant
    Dim i                 As Long
    
    ErrorType = NO_ERROR
    If TradeDate > MaturityDate Then
        ErrorType = XlCVError.xlErrValue
    ElseIf Coupon < 0 Then
        ErrorType = XlCVError.xlErrValue
    ElseIf TradedSpread < 0 Then
        ErrorType = XlCVError.xlErrValue
    ElseIf Recovery < 0 Or Recovery > 1 Then
        ErrorType = XlCVError.xlErrValue
    ElseIf HazardRate < 0 Or Recovery > 1 Then
        ErrorType = XlCVError.xlErrValue
    ElseIf YieldDatesRange.Columns.Count <> 1 Then
        ErrorType = XlCVError.xlErrValue
    ElseIf YieldRatesRange.Columns.Count <> 1 Then
        ErrorType = XlCVError.xlErrValue
    ElseIf YieldRatesRange.Rows.Count <> YieldDatesRange.Rows.Count Then
        ErrorType = XlCVError.xlErrValue
    Else
        YieldDatesArray = YieldDatesRange.Value
        YieldRatesArray = YieldRatesRange.Value
        
        i = 1
        lastDate = Application.WorkDay(TradeDate, SPOT_DAYS)
        
        Do While ErrorType = NO_ERROR And i <= UBound(YieldDatesArray, 1)
            thisDate = YieldDatesArray(i, 1)
            If Not VBA.IsDate(thisDate) Then
                ErrorType = XlCVError.xlErrValue
            ElseIf thisDate <= lastDate Then
                ErrorType = XlCVError.xlErrValue
            ElseIf Not VBA.IsNumeric(YieldRatesArray(i, 1)) Then
                ErrorType = XlCVError.xlErrValue
            End If
            i = i + 1
            lastDate = thisDate
        Loop
    End If
    ValidateInputData = NO_ERROR
End Function

'########################################################################################
'# Load worksheet data into VBA Data Structures
'########################################################################################

Private Function LoadCreditMarket( _
        ByVal CurveDate As Date, _
        ByVal Recovery As Double, _
        Optional ByVal HazardRate As Variant, _
        Optional ByVal TradedSpread As Variant) _
        As CreditMarketData
    
    ' load market data relating to Credit into data structure
    
    Dim Market As CreditMarketData
    
    With Market
        .CurveDate = CurveDate
        .Recovery = Recovery
        If Not IsMissing(HazardRate) Then .HazardRate = HazardRate
        If Not IsMissing(TradedSpread) Then .TradedSpread = TradedSpread
    End With
    
    LoadCreditMarket = Market
    
End Function

Private Function LoadCDSContract( _
        ByVal TradeDate As Date, _
        ByVal MaturityDate As Date, _
        ByVal Notional As Double, _
        ByVal Coupon As Double) _
        As CDSContractSpecification
    
    ' load various contract details into CDS Contract data structure

    Dim Contract As CDSContractSpecification
    
    With Contract
        .TradeDate = TradeDate
        .CashDate = Application.WorksheetFunction.WorkDay(.TradeDate, CASH_DAYS)
        .MaturityDate = MaturityDate
        .Notional = Notional
        .Coupon = Coupon
    End With
    
    LoadCDSContract = Contract
    
End Function

Private Function LoadYieldMarket( _
        ByVal CurveDate As Date, _
        ByRef YieldDatesRange As Range, _
        ByRef YieldRatesRange As Range) _
        As YieldMarketData
    
    ' Loads market data from spreadsheet into a yield curve data structure
    
    Dim YieldDatesArray() As Variant
    Dim YieldRatesArray() As Variant
    Dim i                 As Long
    Dim Market            As YieldMarketData

    YieldDatesArray = YieldDatesRange.Value
    YieldRatesArray = YieldRatesRange.Value
    
    With Market
        .CurveDate = CurveDate
        Set .Dates = New Collection
        Set .Rates = New Collection
        
        For i = 1 To UBound(YieldRatesArray)
            .Dates.Add VBA.CDate(YieldDatesArray(i, 1))
            .Rates.Add VBA.CDate(YieldRatesArray(i, 1))
        Next i
    End With
    LoadYieldMarket = Market
    
End Function

'########################################################################################
'# CDS Valuation
'########################################################################################

Private Function CDSNPV( _
        ByRef CDSContract As CDSContractSpecification, _
        ByRef ValuationDates As CDSValuationDates, _
        ByRef YieldMarket As YieldMarketData, _
        ByRef CreditMarket As CreditMarketData, _
        Optional isDirty As Boolean = True) _
        As Double
    
    ' Calculates Value of a CDS contract given market data, including a credit market where the zero hazard rate is know.
    ' if the traded spread is known, but not the hazard rate, then use CDSNPVfromTradedSpread()
        
    Dim AccruedOnDefault As Double
    Dim ProtectionLeg    As Double
    Dim PremiumLeg       As Double
    Dim CashDF           As Double
    Dim Interest         As Double
    
    CashDF = DiscountFactorFromMarket(CDSContract.CashDate, YieldMarket)
    PremiumLeg = PremiumLegNPV(CDSContract, ValuationDates.PaymentDates, ValuationDates.AccrualDates, YieldMarket, CreditMarket)
    ProtectionLeg = ProtectionLegNPV(CDSContract, ValuationDates.ProtectionNodes, YieldMarket, CreditMarket)
    AccruedOnDefault = AccruedOnDefaultNPV(CDSContract, ValuationDates.AccruedOnDefaultNodes, ValuationDates.AccruedOnDefaultAccrualStartDates, YieldMarket, CreditMarket)
    
    If isDirty Then
        Interest = 0
    Else
        Interest = CDSContract.Notional * CDSContract.Coupon * Act360(ValuationDates.AccrualDates.Item(1), CDSContract.TradeDate)
    End If
    
    CDSNPV = (PremiumLeg + AccruedOnDefault - ProtectionLeg) / CashDF - Interest
End Function


Private Function CDSNPVfromTradedSpread( _
        ByRef CDSContract As CDSContractSpecification, _
        ByRef ValuationDates As CDSValuationDates, _
        ByRef YieldMarket As YieldMarketData, _
        ByRef CreditMarket As CreditMarketData, _
        ByRef isDirty As Boolean) _
        As Double
    
    ' Calculates CDS NPV given a traded spread
    
    CreditMarket = FindHazardRate(CDSContract, ValuationDates, YieldMarket, CreditMarket)
    CDSNPVfromTradedSpread = CDSNPV(CDSContract, ValuationDates, YieldMarket, CreditMarket, isDirty)
        
End Function

Private Function FindHazardRate( _
        ByRef CDSContract As CDSContractSpecification, _
        ByRef ValuationDates As CDSValuationDates, _
        ByRef YieldMarket As YieldMarketData, _
        ByRef CreditMarket As CreditMarketData) _
        As CreditMarketData
    
    ' uses secant root finding to solve for the hazard rate which prices a given CDS at par (i.e. clean pv = 0)

    Const WORKING_NOTIONAL        As Double = 1000000000       ' 1E9
    Const SECOND_GUESS_MULTIPLIER As Double = 1.1
    Const EPSILON                 As Double = 1E-14            ' 1E-14
    
    Dim ImpliedCDSContract  As CDSContractSpecification
    Dim ImpliedCreditMarket As CreditMarketData
    Dim Interest            As Double
    Dim H0                  As Double
    Dim H1                  As Double
    Dim V0                  As Double
    Dim V1                  As Double
    
    ImpliedCDSContract = LoadCDSContract(CDSContract.TradeDate, CDSContract.MaturityDate, WORKING_NOTIONAL, CreditMarket.TradedSpread)
    ImpliedCreditMarket = CreditMarket
    
    H0 = CreditMarket.TradedSpread / (1 - CreditMarket.Recovery)
    
    ImpliedCreditMarket.HazardRate = H0
    V0 = CDSNPV(ImpliedCDSContract, ValuationDates, YieldMarket, ImpliedCreditMarket, False)
    
    H1 = H0 * SECOND_GUESS_MULTIPLIER
    ImpliedCreditMarket.HazardRate = H1
    
    Do Until Abs((H1 - H0) / H0) < EPSILON
        V1 = CDSNPV(ImpliedCDSContract, ValuationDates, YieldMarket, ImpliedCreditMarket, False)
        If (V1 - V0) = 0 Then Exit Do
        ImpliedCreditMarket.HazardRate = H0 + (H0 - H1) * V0 / (V1 - V0)
        H0 = H1
        H1 = ImpliedCreditMarket.HazardRate
        V0 = V1
    Loop
    
    FindHazardRate = ImpliedCreditMarket
End Function

Private Function PremiumLegNPV( _
        ByRef CDSContract As CDSContractSpecification, _
        ByRef PaymentDates As Collection, _
        ByRef AccrualDates As Collection, _
        ByRef YieldMarket As YieldMarketData, _
        ByRef CreditMarket As CreditMarketData) _
        As Double
        
    ' value of the premiums paid up until default
    
    Dim Valuation() As ValuationNode
    Dim i           As Long
    Dim NPV         As Double
    
    ReDim Valuation(1 To PaymentDates.Count)

    Valuation(1).AccrualDate = AccrualDates.Item(1)
    
    For i = 2 To UBound(Valuation)
        With Valuation(i)
            .AccrualDate = AccrualDates.Item(i)
            .PaymentDate = PaymentDates.Item(i)
            .PeriodDCC = Act360(Valuation(i - 1).AccrualDate, .AccrualDate)
            .DiscountFactor = DiscountFactorFromMarket(.PaymentDate, YieldMarket)
            .ProtectionDate = .AccrualDate
            .SurvivalProbability = SurvivalProbability(.ProtectionDate, CreditMarket)
            .PV = CDSContract.Notional * CDSContract.Coupon * .PeriodDCC * .DiscountFactor * .SurvivalProbability
            NPV = NPV + .PV
        End With
    Next i
    PremiumLegNPV = NPV
    
End Function

Private Function ProtectionLegNPV( _
        ByRef CDSContract As CDSContractSpecification, _
        ByRef ProtectionNodes As Collection, _
        ByRef YieldMarket As YieldMarketData, _
        ByRef CreditMarket As CreditMarketData) _
        As Double
    
    ' Value of the protection amount paid upon default
    
    Dim Valuation() As ValuationNode
    Dim i As Long
    Dim NPV As Double
    
    ReDim Valuation(1 To ProtectionNodes.Count)
    
    ' calculated for each node
    
    For i = 1 To UBound(Valuation)
        With Valuation(i)
            .ProtectionDate = ProtectionNodes.Item(i)
            .Zero = ZeroRate(.ProtectionDate, YieldMarket)
            .DiscountFactor = DiscountFactorFromZero(YieldMarket.CurveDate, .ProtectionDate, .Zero)
            .LogDF = -.Zero * Act365(YieldMarket.CurveDate, .ProtectionDate)
            .SurvivalProbability = SurvivalProbability(.ProtectionDate, CreditMarket)
        End With
    Next i
    
    ' calcaulted for each interval between nodes (hence start at 2)
    
    For i = 2 To UBound(Valuation)
        With Valuation(i)
            .Interest = Valuation(i - 1).LogDF - .LogDF
            .Hazard = CreditMarket.HazardRate * Act365(Valuation(i - 1).ProtectionDate, .ProtectionDate)
            .PV = CDSContract.Notional * (1 - CreditMarket.Recovery) * .Hazard / (.Hazard + .Interest) * _
                (Valuation(i - 1).DiscountFactor * Valuation(i - 1).SurvivalProbability - .DiscountFactor * .SurvivalProbability)
            NPV = NPV + .PV
        End With
    Next i
    ProtectionLegNPV = NPV

End Function

Private Function AccruedOnDefaultNPV( _
        ByRef CDSContract As CDSContractSpecification, _
        ByRef Nodes As Collection, _
        ByRef StartDates As Collection, _
        ByRef YieldMarket As YieldMarketData, _
        ByRef CreditMarket As CreditMarketData) _
        As Double
    
    ' Value of the accrued interest paid on default
    
    Const ADJUSTEMENT_FOR_CONTINUOUS_TIME As Double = 0.5
    ' accrual following default is rounded up to nearest full day, continuous approximation out by half day on average
    
    Dim Valuation() As ValuationNode
    Dim i As Variant
    Dim NPV As Double
    
    ReDim Valuation(1 To Nodes.Count)
    
    ' calculated for each node
    For i = 1 To UBound(Valuation)
        With Valuation(i)
            .ProtectionDate = Nodes.Item(i)
            .Zero = ZeroRate(.ProtectionDate, YieldMarket)
            .DiscountFactor = DiscountFactorFromZero(YieldMarket.CurveDate, .ProtectionDate, .Zero)
            .LogDF = -.Zero * Act365(YieldMarket.CurveDate, .ProtectionDate)
            .SurvivalProbability = SurvivalProbability(.ProtectionDate, CreditMarket)
            .RiskyDF = .DiscountFactor * .SurvivalProbability
            .AccrualStart = StartDates.Item(i)
        End With
    Next i
    
    ' calculated for each interval between nodes (hence start at 2)
    For i = 2 To Nodes.Count
        With Valuation(i)
            .StartDCC = Act360(.AccrualStart, Valuation(i - 1).ProtectionDate + ADJUSTEMENT_FOR_CONTINUOUS_TIME)
            .PeriodDCC = Act360(Valuation(i - 1).ProtectionDate, .ProtectionDate)
            .Interest = Valuation(i - 1).LogDF - .LogDF
            .Hazard = CreditMarket.HazardRate * Act365(Valuation(i - 1).ProtectionDate, .ProtectionDate)
            .RiskyDFChange = Valuation(i - 1).RiskyDF - .RiskyDF
            .PV = (.RiskyDFChange * (.PeriodDCC / (.Hazard + .Interest) + .StartDCC) - .PeriodDCC * .RiskyDF) * _
                  .Hazard / (.Hazard + .Interest) * CDSContract.Notional * CDSContract.Coupon
            NPV = NPV + .PV
        End With
    Next i

    AccruedOnDefaultNPV = NPV
End Function

'########################################################################################
'# Market Data Functions
'########################################################################################

Private Function DiscountFactorFromMarket(ByVal Target As Date, ByRef Market As YieldMarketData) As Double
    ' Discount Factor for a target date using the yield curve given
    DiscountFactorFromMarket = DiscountFactorFromZero(Market.CurveDate, Target, ZeroRate(Target, Market))
End Function

Private Function DiscountFactorFromZero(ByVal CurveDate As Date, ByVal Target As Date, ByVal Zero As Double) As Double
    ' Discount Factor for a target date using the zero-coupon interest rate given
    ' Interest rate is continuously compounded zero rate act/365
    DiscountFactorFromZero = Exp(-Act365(CurveDate, Target) * Zero)
End Function

Private Function ZeroRate( _
        ByVal Target As Date, _
        ByRef Market As YieldMarketData) _
        As Double
    ' Find the zero rate for a target date by interpolation along a given yield curve
    ' interpolation is linear in zt
    ' rates are continously compounded act/365
    
    Dim i As Long
    Dim t  As Double
    Dim t1 As Double, t2 As Double
    Dim z1 As Double, z2 As Double
    
    With Market
    
        If Target < .Dates.Item(1) Or .Dates.Count = 1 Then
            ' times before the first node have flat zero rate, just return the first zero rate
            ZeroRate = .Rates.Item(1)
        Else
            If Target >= .Dates.Item(.Dates.Count) Then
                ' times after the last node are extrapolated from the last two nodes
                i = .Dates.Count - 1
            Else
                ' linear search for the first node to interpolate between
                For i = 1 To .Dates.Count - 1
                    If Target < .Dates.Item(i + 1) Then Exit For
                Next i
            End If
            
            ' linear interpolation/extrapolation in zt
            t1 = .Dates.Item(i) - Market.CurveDate
            t2 = .Dates.Item(i + 1) - Market.CurveDate
            z1 = .Rates.Item(i)
            z2 = .Rates.Item(i + 1)
            t = Target - .CurveDate
            ZeroRate = (z1 * t1 + (t - t1) / (t2 - t1) * (z2 * t2 - z1 * t1)) / t
        End If
        
    End With
        
End Function

Private Function SurvivalProbability( _
        ByVal Target As Date, _
        ByRef Market As CreditMarketData) _
        As Double
    ' Survival probability for a target date using some hazard rate
    ' Hazard rate is continuously compounded zero rate act/365
    
    SurvivalProbability = Exp(-Market.HazardRate * Act365(Market.CurveDate, Target))
End Function

'########################################################################################
'# Date Functions
'########################################################################################

Private Function ModFol(ByVal UnadjustedDate As Date) As Date
    ' Roll date forward to next good business day
    ' unless that would result in a new month,
    ' in which case roll backwards
    Dim Forward As Date
    Forward = Application.WorksheetFunction.WorkDay(UnadjustedDate - 1, 1)
    If VBA.Month(Forward) = VBA.Month(UnadjustedDate) Then
        ModFol = Forward
    Else
        ModFol = Application.WorksheetFunction.WorkDay(UnadjustedDate + 1, -1)
    End If
End Function

Private Function Act360(ByVal StartDate As Date, ByVal EndDate As Date) As Double
    ' Year fraction using Act/360 convention
    Const DAYS_IN_YEAR_360 As Long = 360
    Act360 = (EndDate - StartDate) / DAYS_IN_YEAR_360
End Function

Private Function Act365(ByVal StartDate As Date, ByVal EndDate As Date) As Double
    ' Year fraction using Act/365 convention
    Const DAYS_IN_YEAR_365 As Long = 365
    Act365 = (EndDate - StartDate) / DAYS_IN_YEAR_365
End Function


Private Function GenerateValuationDates(ByRef CDSContract As CDSContractSpecification, ByRef YieldMarket As YieldMarketData) As CDSValuationDates
    Dim ValuationDates As CDSValuationDates
    ' generate all dates needed to value a cds contract
    
    GeneratePremiumLegDates ValuationDates, CDSContract
    GenerateProtectionLegDates ValuationDates, CDSContract, YieldMarket
    GenerateAccruedOnDefaultDates ValuationDates
    
    GenerateValuationDates = ValuationDates
    
End Function

Private Sub GeneratePremiumLegDates( _
        ByRef ValuationDates As CDSValuationDates, _
        ByRef CDSContract As CDSContractSpecification)

    ' Generate Dates for premium accrual and payments
    
    Dim Unadjusted As Date
    Dim Adjusted   As Date
    
    With ValuationDates
        Set .PaymentDates = New Collection
        Set .AccrualDates = New Collection
        
        Unadjusted = CDSContract.MaturityDate
        Adjusted = ModFol(Unadjusted)
        
        .PaymentDates.Add Adjusted
        .AccrualDates.Add Unadjusted
        
        Do While Adjusted > CDSContract.TradeDate
            Unadjusted = Application.WorksheetFunction.CoupPcd(Unadjusted - 1, CDSContract.MaturityDate, QUARTERS_IN_YEAR)
            Adjusted = ModFol(Unadjusted)
            .PaymentDates.Add Item:=Adjusted, before:=1
            .AccrualDates.Add Item:=Adjusted - 1, before:=1
        Loop
    End With
End Sub

Private Sub GenerateProtectionLegDates( _
        ByRef ValuationDates As CDSValuationDates, _
        ByRef CDSContract As CDSContractSpecification, _
        ByRef YieldMarket As YieldMarketData)
        
    ' Generate nodes for protection leg integration
    ' between each node there is flat forward interest rates and flat forward hazard rates
        
    Dim Node As Variant
        
    With ValuationDates
    
        Set .ProtectionNodes = New Collection
        
        .ProtectionNodes.Add CDSContract.TradeDate
        
        For Each Node In YieldMarket.Dates
            If Node >= CDSContract.MaturityDate Then
                Exit For
            Else
                .ProtectionNodes.Add Node
            End If
        Next Node
        
        .ProtectionNodes.Add CDSContract.MaturityDate
    End With
End Sub

Private Sub GenerateAccruedOnDefaultDates(ByRef ValuationDates As CDSValuationDates)
    
    ' Combine dates for premium and protection legs
    ' between each node there are flat forward interest rates, flat forward hazard rates, and linear accrual yearfractions
    
    Dim i As Long, j As Long
    
    With ValuationDates
        Set .AccruedOnDefaultNodes = New Collection
        Set .AccruedOnDefaultAccrualStartDates = New Collection
        
        i = 2 ' ignore first accrual date - is before the trade date
        j = 1
        
        Do Until i > .AccrualDates.Count And j > .ProtectionNodes.Count
            If i > .AccrualDates.Count Then
                ' no more accrual dates, add rest of the protection dates
                For j = j To .ProtectionNodes.Count
                    .AccruedOnDefaultNodes.Add .ProtectionNodes.Item(j)
                    .AccruedOnDefaultAccrualStartDates.Add .AccrualDates.Item(i - 1)
                Next j
            ElseIf j > .ProtectionNodes.Count Then
                ' no more protection dates, add rest of the accrual dates
                For i = i To .AccrualDates.Count
                    .AccruedOnDefaultNodes.Add .AccrualDates.Item(i)
                    .AccruedOnDefaultAccrualStartDates.Add .AccrualDates.Item(i - 1)
                Next i
            ElseIf .AccrualDates.Item(i) < .ProtectionNodes.Item(j) Then
                .AccruedOnDefaultNodes.Add .AccrualDates.Item(i)
                .AccruedOnDefaultAccrualStartDates.Add .AccrualDates.Item(i - 1)
                i = i + 1
            ElseIf .AccrualDates.Item(i) = .ProtectionNodes.Item(j) Then
                .AccruedOnDefaultNodes.Add .AccrualDates.Item(i)
                .AccruedOnDefaultAccrualStartDates.Add .AccrualDates.Item(i - 1)
                i = i + 1
                j = j + 1
            ElseIf .AccrualDates.Item(i) > .ProtectionNodes.Item(j) Then
                .AccruedOnDefaultNodes.Add .ProtectionNodes.Item(j)
                .AccruedOnDefaultAccrualStartDates.Add .AccrualDates.Item(i - 1)
                j = j + 1
            End If
        Loop
        
    End With
    
End Sub


Public Module FramVar

   '- Display Variables
   Public DevWidth, DevHeight As Integer
   Public FormWidth, FormHeight As Integer

   '- Common Loop Variables
   Public FramVersion As String
   Public Stk, Age, Fish, TStep, PTerm, Term, TermStat As Integer
   Public RecordsetSelectionType As Integer
   Public SkipJim As Integer
   Public File_Name As String
   Public VersionNumberChanges() As String

   '- Database Variables 
   Public FVSdatabasename As String
   Public FVSshortname As String
   Public FVSdatabasepath As String
   Public TAMMSpreadSheet As String
   Public TAMMSpreadSheetPath As String
   Public BACKFRAMSpreadSheet As String
   Public BACKFRAMSpreadSheetPath As String
   Public FramDataSet As New System.Data.DataSet()
   Public TransferDataSet As New System.Data.DataSet()
   Public FramDB As New OleDb.OleDbConnection
    Public TransDB As New OleDb.OleDbConnection
    Public TransBP As New OleDb.OleDbConnection
   Public StockRecruitDA As New System.Data.OleDb.OleDbDataAdapter

   '- Run Selection Variables
   Public RunIDSelect As Integer
   Public RunIDDelete As Integer
   Public RunIDTransfer() As Integer
   Public NumTransferID As Integer
   Public NewRunID As Integer
   Public RunIDNameSelect As String
   Public RunIDTitleSelect As String
   Public RunIDCommentsSelect As String
   Public RunIDCreationDateSelect As DateTime
   Public RunIDModifyInputDateSelect As DateTime
    Public RunIDRunTimeDateSelect As DateTime
    Public RunIDYearSelect As String
    Public RunIDTypeSelect As String
   Public OptionReplaceQuota As Boolean
   Public OptionOldTAMMformat As Boolean
   Public OptionUseTAMMfws As Boolean
   Public OptionChinookBYAEQ As Integer
   Public ModelRunBPSelect As Boolean
    Public SelectSpeciesName As String
    Public SizeLimitFix As Boolean
    Public SizeLimitOnly As Boolean
   Public RunIDmultiDelete() As Integer 'Pete 12/13 variable for multi-run deletion code
   Public multiRunPass As String 'Pete 12/13 variable for passing multi-run deletion during loop
    Public multiRunDeleteMode As Boolean 'Pete 12/13 variable for bypassing some content in delete mode
    Public TAMMName As String
    Public CoastalIter As String
    Public FRAMVers As String


   '- Input Edit Change Variables
   Public ChangeAnyInput As Boolean
   Public ChangeBackFram As Boolean
   Public ChangeFishScalers As Boolean
   Public ChangeNonRetention As Boolean
   Public ChangePSCMaxER As Boolean
   Public ChangeSizeLimit As Boolean
   Public ChangeStockFishScaler As Boolean
    Public ChangeStockRecruit As Boolean
    Public AnyChange As Boolean

   '- Base Period Variables
   Public BasePeriodID As Integer
   Public BasePeriodIDSelect As Integer
    Public BasePeriodName As String
    Public BPSL_No_ID As Boolean
   Public SpeciesName As String
   Public NumStk As Integer
   Public NumFish As Integer
   Public NumSteps As Integer
   Public NumAge As Integer
   Public MinAge As Integer
   Public MaxAge As Integer
   Public BasePeriodDate As Date
   Public BasePeriodComments As String
   Public StockVersion As Integer
   Public FisheryVersion As Integer
   Public TimeStepVersion As Integer
   Public BaseCohortSize(,) As Double
   Public BaseExploitationRate(,,,) As Double
   Public AnyBaseRate(,) As Integer
   Public BaseSubLegalRate(,,,) As Double
   Public MaturationRate(,,) As Double
   Public ShakerMortRate(,) As Double
   Public AEQ(,,) As Double
    Public EncounterRateAdjustment(,,) As Double
    Public MSFBaseCatch As Double
    Public MSFBaseSubEncounters As Double
    Public MSFNewSubEncounters As Double
    Public MSFBaseShakers As Double
    Public MSFNewShakers As Double
    Public MSFSubEncDiff As Double
    Public MSFLegalEncounters(,,,) As Double
    Public MSFTotalEncounters(,) As Double
    Public MSFNewCatch As Double
    Public MSFBaseLegalEncounters

   '- FramCalcs Variables
   Public LandedCatch(,,,) As Double
    Public NonRetention(,,,) As Double
    Public NRLegal(,,,,) As Double
   Public Shakers(,,,) As Double
    Public LegalShakers(76, 5, 73, 4) As Double
   Public DropOff(,,,) As Double
   Public MSFLandedCatch(,,,) As Double
   Public MSFNonRetention(,,,) As Double
   Public MSFShakers(,,,) As Double
   Public MSFDropOff(,,,) As Double
   Public Encounters(,,,) As Double
   Public MSFEncounters(,,,) As Double
   Public Cohort(,,,) As Double
    Public Escape(,,) As Double
    Public NSFQuotaTotal(,) As Double
    Public MSFQuotaTotal(,) As Double
    Public Quotacatch As Double
    Public MSFQuotaEncounters As Double
    Public T4CohortFlag As Boolean
    Public T4CohortFlag2 As Boolean

   '- Base Stock Variables
   Public StockID() As Integer
   Public ProductionRegion() As Integer
   Public ManagementUnit() As Integer
   Public StockName() As String
   Public StockTitle() As String
   Public VonBertL(,), VonBertK(,), VonBertT(,), VonBertCV(,,) As Double
   Public MidTimeStep() As Double

   '- Base Fishery Variables
   Public FisheryID() As Integer
   Public FisheryName() As String
   Public FisheryTitle() As String
   Public ModelStockProportion() As Double
   Public NaturalMortality(,) As Double
   Public IncidentalRate(,) As Double
   Public TerminalFisheryFlag(,) As Integer
   Public ChinookBaseEncounterAdjustment(,) As Double
   Public ChinookBaseSizeLimit(,) As Integer
   Public ChinookBaseLegProp As Boolean
   Public BaseLegalProportion, BaseSubLegalProportion As Double

   '- Base Time Step Variables
   Public TimeStepID() As Integer
   Public TimeStepName() As String
   Public TimeStepTitle() As String

    '- Run Input Variables
    Public AutoSizeLimitCalcs As Boolean
    Public StockRecruit(,,) As Double
    Public FisheryComment(,)
   Public FisheryScaler(,) As Double
    Public FisheryQuota(,) As Double
    Public FisheryQuotaCompare(NumFish, NumSteps) As Double
    Public FisheryFlag(,) As Integer
    Public NonRetentionComment(,)
   Public NonRetentionFlag(,) As Integer
   Public NonRetentionInput(,,) As Double
   Public MSFFisheryScaler(,) As Double
   Public MSFFisheryQuota(,) As Double
   Public MarkSelectiveMortRate(,) As Double
   Public MarkSelectiveMarkMisID(,) As Double
   Public MarkSelectiveUnMarkMisID(,) As Double
   Public MarkSelectiveIncRate(,) As Double
    Public MinSizeLimit(,) As Integer
    Public NewSizeLimit(,) As Integer
   Public MaxSizeLimit(,) As Integer
   Public StockFishRateScalers(,,) As Double
   Public RunTAMMIter As Integer
   Public RunBackFramFlag As Integer
    Public PSCMaxER() As Double
    Public KeepIter As Boolean
    Public Itercount As Integer

   '- Run Calculation Variables
   Public TotalLandedCatch(,) As Double
    Public TotalNonRetention(,) As Double
    Public FTNonRetention(,) As Double
   Public TotalEncounters(,) As Double
    Public TotalLegalShakers(73, 4) As Double
   Public TotalShakers(,) As Double
   Public TotalDropOff(,) As Double
   Public PropLegCatch(,) As Double
   Public PropSubPop(,) As Double
   Public CNRShakers(,) As Double
   Public LegalProportion, SubLegalProportion As Double
   Public LegalPop, SubLegalPop As Double
   Public AnyNegativeEscapement, NegativeEsc(,) As Integer
   Public TammTransferSave As Boolean
    Public MSFBiasFlag As Boolean
    Public NoMSFBiasCalcs As Boolean
    Public MessageFlag As Boolean = False
    Public msgFlag As Boolean = False
   '###################################################Pete-12/17/12.
    Public ERgtrOne(,) As Boolean 'This is associated with the MSF bias correction calculations
   '###################################################Pete-12/17/12.

   '=====================================================================
   'Pete 12/13 Pete External Sublegals Variables
   Public RunEncounterRateAdjustment(,,) As Double
   Public TargetRatio(,,) As Double
    Public Kfat(,,), Kfat2(,,) As Double 'Temporary in-update variable for computing new RunEncounterRateAdjustment
   Public UpdBy(,,) As String
   Public UpdWhen(,,) As DateTime
   Public UpdateRunEncounterRateAdjustment As Boolean
   Public WhoUpdated As String
   Public dsSLquery As New DataSet
   Public FinalUpdatePass As Boolean
   Public ReadOldCmd As Boolean
   '=====================================================================


   ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
   ''#################### BEGIN NEW CODE BLOCK ###############################################################
   ''These are all variables associated with modeling alternative
   ''size limits in Puget Sound sport fisheries
   'Public SizeLimitScenario As Boolean
   'Public AltFlag(,) As Double
   'Public AltLimitNS(,), AltLimitMSF(,) As Double
   'Public ShakerFlagNS(,), ShakerFlagMSF(,) As Double
   'Public LSRatioNS(,), LSRatioMSF(,) As Double
   'Public ExtShakerNS(,), ExtShakerMSF(,) As Double
   'Public NSShakerExtTotal(,), MSFShakerExtTotal(,) As Double
   'Public MSFEncountersTotal(,), NSEncountersTotal(,) As Double 'This variable is required to store total legal encounters (currently added to fishery total enc)
   'Public NSLegalProp, NSSublegalProp, MSFLegalProp, MSFSublegalProp, CNRLegalProp, CNRSublegalProp As Double
   'Public ExternalBaseRatio(,) As Double 'This is where the external base Sublegal:Legal Ratio Goes
   'Public RescaleFactor(,) As Double 'This is the multiplier to rescale legal ecnounters (or LC)
   ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
   ''#################### END CODE BLOCK #################################################################



   '- Chinook TAMM Variables
   Public TammCatch(,) As Double
   Public TammEscape(,) As Double
   Public TammEstimate(,) As Double
   Public TammTermRun() As Double
   Public TammPSER(,) As Double
   Public TammScaler(,) As Double
   Public TammIteration As Integer
   Public TammChinookConverge As Integer
   Public TotalChinEsc(,) As Double
   Public NewStockFishRateScalers As Integer
   Public TammChinookRunFlag As Integer
   Public TamkFish() As Integer
   Public SpsYrSpl As Double
   Public TSkFWSpt!, TSkMSA!, TSnFWSpt!, TSnMSA!
    Public TNkFWSpt!, TNkMSA!, THCFWSpt!
    

   '- Coho TAMM Variables
   Public CohoTammRate(,) As Double
   Public CohoTammFlag(,) As Integer
   Public CohoTammFish(,) As Integer
   Public CohoTime4Cohort() As Double
   Public TaaEtrsNum() As Integer
   Public TaaEtrsStk(,) As Integer
   Public TaaEtrsFish(,) As Integer
   Public TaaEtrsTStep(,) As Integer
   Public TaaEtrsType() As Integer
   Public TaaEtrsName() As String
   Public SaveTermFlag(,) As Integer
    Public SaveTermQuota(,) As Double
    Public SaveCoastalQuota(,) As Double

    '- Backwards FRAM
    Public AgeTSCatch(,,) As Double
    Public AgeTSCatchTerm(,,) As Double
    Public SumTSCatch(,) As Double
    Public BackwardsComment()
    Public BackwardsFlag(Stk + 1) As Integer
   Public BackwardsFRAMFlag As Integer
   Public BackwardsTarget() As Double
   Public BackwardsChinook(,) As Double

   Public BFYearSelection As Integer
   Public BFYearSelectType As Integer
   Public BFEscYears() As Integer
   Public BFCatchYears() As Integer
   Public BackScaler(,) As Double
   Public BackEsc(,) As Double
   Public BackChinScaler(,,) As Double
    Public BackChinEsc(,,) As Double
    Public BackFRAMIteration As Integer
    Public BkMethod As Integer
    Public ERBKMethod(,,) As Double
    Public EscDiffArray(,,) As Double
    Public FirstIter As Integer
    Public InitialCohort() As Double
    Public InitialCohortM() As Double
    Public TermChinRun(,) As Double
    Public DoneIterating As Integer
    Public MatRateCounter As Integer
    Public NumChinTermRuns As Integer
    Public OldScalar(,,) As Double
    Public RunBackwardsTarget() As Double
    Public RunBackwardsFlag() As Integer
    Public TempCohort(,) As Double
   Public TermRunName() As String
   Public TermRunStock(), TermStockNum(), TFish(,), TTime(,) As Integer
    Public ChinSurvMult() As Double
    Public StartRate(,) As Double
    Public SurvMultSp() As Double
    Public SaveInitialFlag As Boolean
    Public BackFramSave As Boolean
    Public TermCounter As Integer
    Public TRun As Integer
    Public xvar As Integer


   '- Brood Year Chinook FRAM
   Public BYERFlag As Integer
   Public BYCohort(,,,,) As Double
   Public BYEscape(,,,) As Double
   Public BYLandedCatch(,,,,) As Double
   Public BYNonRetention(,,,,) As Double
   Public BYShakers(,,,,) As Double
   Public BYDropOff(,,,,) As Double
   Public BYMSFLandedCatch(,,,,) As Double
   Public BYMSFNonRetention(,,,,) As Double
   Public BYMSFShakers(,,,,) As Double
    Public BYMSFDropOff(,,,,) As Double
    Public BY As Integer
    Public Cost As Integer
    Public BYAge As Integer



   '- FRAM Utilities
   Public OldCMDFile, OldCMDFilePath As String
   Public OldOUTFile, OldOUTFilePath As String
    Public Jim As Integer
    Public TransferBPName As String
    Public NewTransferBP As String
    Public ImportBP As Boolean
    Public ImportStock As Boolean
    Public ImportFish As Boolean
    Public ImportTS As Boolean
    Public CoastalIterations As Boolean

   '- Output Reports
   Public CallingRoutine As Integer
   Public ScreenReportType As Integer
   Public DriverSelectionType As Integer
   Public ReportNumber As Integer
   Public NumDriverReports As Integer
   Public MortalityType As Integer
   Public StockSelection() As Integer
   Public NumSelectedStocks As Integer
   Public StockGroupName As String
   Public StockSelectionType As Integer
   Public FisherySelectionType As Integer
   Public FisheryEditSelection As Integer
   Public FisherySelection() As Integer
   Public NumSelectedFisheries As Integer
   Public FisheryGroupName As String
   Public OldDriverFileName As String
   Public TimeStepSelection1 As Integer
   Public TimeStepSelection2 As Integer
   Public TermRunTypeSelection As Integer
   Public TerminalRunReportSelected As Boolean
   Public TermRunBYAEQ As Boolean
   Public ReportFileName As String
   Public ReportDriverName As String
   Public RepStks(,), NumRepStks() As Integer
   Public RepFish(,), NumRepFish() As Integer
   Public RepTStep(,), NumRepGrps As Integer
   Public RepGrpName(), RepGrpType() As String
   '- Multi Group Fishery Variables
   Public SelectFishery, SelectFisheryList(,), FisheryCheckList() As Integer
   Public NumGroupFisheries, NumFisheryGroups As Integer
   Public SelectFisheryName, FisheryGroupNames() As String

   '- Screen Scaling Variables
   Public FormWidthScaler As Double
   Public FVS_BackwardsFram_ReSize As Boolean
   Public FVS_BackwardsResults_ReSize As Boolean
   Public FVS_BackwardsTarget_ReSize As Boolean
   Public FVS_BackwardsYearSelect_ReSize As Boolean
   Public FVS_BasePeriodSelect_ReSize As Boolean
   Public FVS_EditRecordSetInfo_ReSize As Boolean
   Public FVS_FisheryMultiGroupSelect_ReSize As Boolean
   Public FVS_FisheryScalerEdit_ReSize As Boolean
   Public FVS_FisheryScalerScreen_ReSize As Boolean
   Public FVS_FisherySelect_ReSize As Boolean
   Public FVS_FishStkCompScreen_ReSize As Boolean
   Public FVS_FramUtils_ReSize As Boolean
   Public FVS_InputMenu_ReSize As Boolean
   Public FVS_MainMenu_ReSize As Boolean
   Public FVS_ModelRunSelection_ReSize As Boolean
   Public FVS_MortalityReport_ReSize As Boolean
   Public FVS_MortalityTypeSelection_ReSize As Boolean
   Public FVS_NonRetentionEdit_ReSize As Boolean
   Public FVS_Output_ReSize As Boolean
   Public FVS_OutputDriver_ReSize As Boolean
   Public FVS_OutputDriverSelection_ReSize As Boolean
   Public FVS_PSCCohoERScreen_ReSize As Boolean
   Public FVS_PSCMaxER_ReSize As Boolean
   Public FVS_ReportSelection_ReSize As Boolean
   Public FVS_RunModel_ReSize As Boolean
   Public FVS_SaveModelRunInputs_ReSize As Boolean
   Public FVS_ScreenReports_ReSize As Boolean
   Public FVS_SelectiveFisheryScreen_ReSize As Boolean
   Public FVS_SizeLimitEdit_ReSize As Boolean
   Public FVS_StockFisheryScalerEdit_ReSize As Boolean
   Public FVS_StockImpactsPer1000Screen_ReSize As Boolean
   Public FVS_StockRecruitEdit_ReSize As Boolean
   Public FVS_StockSelect_ReSize As Boolean
   Public FVS_PopStatScreen_ReSize As Boolean
    Public FVS_Coweeman_ReSize As Boolean
    Public FVS_ActiveRateScaler_ReSize As Boolean

End Module

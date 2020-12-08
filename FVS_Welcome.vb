Option Explicit Off
Option Strict Off

Public Class FVS_Welcome

   Private Sub FVS_Welcome_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        FramVersion = "2.21Dec15"
      VersionLabel.Text = "Version " & FramVersion
      ReDim VersionNumberChanges(100)
      For Stk As Integer = 0 To 100
         VersionNumberChanges(Stk) = ""
      Next

      '-FramVS Program Development Environment Screen
      DevHeight = 1024
      DevWidth = 1280
      FormHeight = 665
      FormWidth = 827

      If My.Computer.Screen.Bounds.Height = 1024 And My.Computer.Screen.Bounds.Width = 1280 Then
         Call StopFormReSize()
      Else
         Call StartFormReSize()
         Me.Height = FormHeight / (DevHeight / My.Computer.Screen.Bounds.Height)
         Me.Width = FormWidth / (DevWidth / My.Computer.Screen.Bounds.Width)
         Resize_Form(Me)
      End If

      '-FramVS Version Numbers
      VersionNumberChanges(0) = "0.001 Initial Version 2010"
      VersionNumberChanges(1) = "0.002 Chinook FW Sport now included in TAA calculations"
      VersionNumberChanges(2) = "0.003 Fixed some Report Outputs"
      VersionNumberChanges(3) = "0.004 Added Back-Fram-Target Coho 'Load Catch' option"
      VersionNumberChanges(4) = "0.005 Coho PSCTable2.DRV Sort (for TAMM Transfer)"
      VersionNumberChanges(5) = "0.006 Coho TAMM Transfer 'Cut' cell addressing problem"
      VersionNumberChanges(6) = "0.007 TermRunRep fix and Coho TAMM overwrite FisheryFlag and FisheryQuota"
      VersionNumberChanges(7) = "0.008 Many fixes for PSC Coho Tech"
      VersionNumberChanges(8) = "0.009 MSF Parameter Save & PSC Coho Report"
      VersionNumberChanges(9) = "0.010 Fixed Chinook TAMM Transfer per AHB - NonLocals not added to ETRS"
      '        Also changed Target CPU to x86 so program will run on 64-bit machines (MS Jet OLEDB problem)
      VersionNumberChanges(10) = "0.011 Changed RunID Array limits from 50 to 150"

      VersionNumberChanges(11) = "1.00 PFMC Version-1 for PFMC Mar 2011"
      VersionNumberChanges(12) = "1.01 Fixed Coweeman spreadsheet errors (code and ss)"
      VersionNumberChanges(13) = "1.02 Fixed Brood Year CNR calcs"
      VersionNumberChanges(14) = "1.03 PFMC March Changes (BYER Chinook and Coho CNR Load)"
      VersionNumberChanges(15) = "2.00 Major Change in FramCalcs - Retention and MSF can now be run in same Fishery/Time Step"
      '      Changes to Database format for FisheryScalers, Mortality, FisheryMortality and for MS Jet OLEDB
      VersionNumberChanges(16) = "2.01 Fixed Fishery Scaler Screen ClipboardCopy (exception error for *****)"
      VersionNumberChanges(17) = "2.02 Fixed NewModelRunTransfer.Mdb Structure, renamed to NewModelRunTransfer2.Mdb Plus Error Msgs"
      VersionNumberChanges(18) = "2.03-MEW Coho MSF Bias Correction Oct 2011 - Not Impemented for 2012"
      VersionNumberChanges(19) = "2.04-MEW Check for Bad Stock & Fishery Numbers in DRV plus new labels in Quota Screen for MSF"
      VersionNumberChanges(20) = "2.05-MEW Fixed Cohort Array Zero problem (5 time step hard code)"
      VersionNumberChanges(21) = "2.06-MEW Chinook MSF Input Template Fill Spreadsheet included ZERO flagged fisheries"
      VersionNumberChanges(22) = "2.07-PFMC Chinook TAMX fixed - New Dates for Reports - Transfer Routine Fixed - ER Rep Fix"
      VersionNumberChanges(23) = "2.08-Coweeman Spreadsheet Fix, Time/Date for Screen Report Copy"
      VersionNumberChanges(24) = "2.09-Additional Coweeman Spreadsheet Automation for New Coweeman TEMPLATE " ' "2.09"
      VersionNumberChanges(25) = "2.10-Chinook Backwards FRAM Terminal Run Read Capability added" ' "2.10"
      VersionNumberChanges(26) = "2.11-Coho MSF Bias Correction and Chinook Pop Stat Code Added" ' "2.11"
      VersionNumberChanges(28) = "2.12-Size Limit Fix, External Sublegal Algorithms, Multi-Run Delete Functionality, AutoScroll on all forms,"
      VersionNumberChanges(27) = "  (2.12 cont.) Check for BaseIDs w/ same BPname in TransferDBs, RunTammIter=0 added, TS4 Recycling of 3s Added for Col R Sp/Su Stks" ' "2.12"
      VersionNumberChanges(29) = "2.13-Chinook Backwards FRAM fix to achieve exact terminal run target for all stocks" ' "2.13"
      VersionNumberChanges(30) = "  (2.14 cont.) - fixed copy/paste error on MSF inputs; fixed Coho bias flag issue (Jan 2015)" ' "2.14 AHB Oct-2014
      VersionNumberChanges(31) = "2.14 - Added Functionality to Export (and Import) ACCESS Base Period Data; " ' "2.14 AHB Oct-2014
        VersionNumberChanges(32) = "(2.14 cont.)  added TRun column to export TRS including non-landed mortality into Coho TAMM - AHB Dec-2014"
        VersionNumberChanges(32) = "(2.14 cont.) added button on RUN SCREEN to revert to old-style handling of cohort aging for stks with abundance of zero" 'AHB Feb-2015
        VersionNumberChanges(32) = "2.15 - 78 stock version with option for new midpoint of growth functions. "
        VersionNumberChanges(32) = "(2.15 cont.)Eliminated White from SPS ETRS TermChinAbun(4) - no effect on TAMM calcs. Elimimated accounting for 13+ twice in TRS calc" 'AHB Nov/Dec - 2015
        VersionNumberChanges(32) = "(2.15 cont.)Added base period transfer functionality for coho" 'AHB Jan 2016
        VersionNumberChanges(32) = "(2.15 cont.)Changed coho bias correction flag to only needing to be checked when running without bias correction" 'AHB Feb 2016
        VersionNumberChanges(32) = "(2.15 cont.)Added check box to run without bias correction in coho bkFRAM. Redimensioned variable tracking ER > 1 for bc runs" 'AHB Feb 2016
        VersionNumberChanges(32) = "(2.15 cont.)Added MessageBox alerting user of bug with saving BkFRAM scalars when pressing SAVE NEW RECORDSET" 'AHB 2/25/16"
        VersionNumberChanges(33) = "(2.16)Made changes to accommodate new base period stock (MOC); most changes occured in bkFRAM array definitions" 'AHB Feb 2016"
        VersionNumberChanges(33) = "(2.16)Corrected FRAMCalcs/ReadChinookTAMM for fisheries 39,40,49,50 to not flag input as scalar if TAMM input >1 (instead of >10)" 'AHB 8/17/16"
        VersionNumberChanges(33) = "(2.16)Corrected TChinSFTran to treat TAMI input as a catch if >=1 instead of >1 (would treat 1 as 100% Rate)" 'AHB Aug 2016
        VersionNumberChanges(34) = "(2.17)Rename Load buttons to IMPORT and Fill buttons to EXPORT" 'AHB Jan 2017"
        VersionNumberChanges(34) = "(2.17)Automate coastal coho iterations by reading values from TAMM and iterating until convergence is achieved 'AHB Jan 2017"
        VersionNumberChanges(34) = "(2.17)For Size Limit Corrected Chinook Model Runs add SLC to run title  'AHB Jan 2017"
        VersionNumberChanges(34) = "(2.17) Remove formatting changes when loading FRAM input into Excel Input Template. Keep template formatting. Comment out lines in Sub LoadSheetButton & FVS_NonRetentionEdit AHB 2/13/17"
        VersionNumberChanges(34) = "(2.17) standardize formatting and handling of decimal places for quotas, scalara, and NR; add a RunYear to RunID table"
        VersionNumberChanges(34) = "(2.17) update age 2 from 3 constants for the new base period in FRAMUtils\btn_Chin2s3s_Click "
        VersionNumberChanges(34) = "(2.17) prevent TAMM from overwriting modeling of Nooksack Earlies in B'ham Bay net; use BPER instead" 'AHB 3/15/17
        VersionNumberChanges(35) = "(2.18) add comment columns to FishScalers and BkFRAM tables updated transfer routines to import and export new columns'AHB 4/5/17"
        VersionNumberChanges(35) = "(2.18) re-code BkFRAM algorithms for Chin & Coho as descibed in 'BkFRAMAug4_2017.docx' corrected age processing in MortAgeReport (Report5) for Chinook (set to age 3 instead of age 3-5)"
        VersionNumberChanges(35) = "(2.18) clairified functionality and re-labeled 'Save New Recordset' button to 'Save BK_Targets'on bkFRAM screen."
        VersionNumberChanges(35) = "(2.18) checkbox to make age 2 from 3 recruit scalars optional during bkFRAM" 'AHB 8/16/2017
        VersionNumberChanges(35) = "(2.18) made overwriting modeling NooksackEarlies in B'hamNet with TAMM a function of Flag-88, if modeled as a rate it will still overwrite for backwards compatibility" 'AHB 8/18/17
        VersionNumberChanges(36) = "(2.19) fixed Coho bkFRAM convergence issues"
        VersionNumberChanges(37) = "(2.19b) added Pass 1 Pass 2 automation; automatic SubLegal updating and size limit correction with each model run 3/5/19"
        VersionNumberChanges(38) = "(2.20) see updates described in  - Proposed FRAM Changes for 2020 Pre-season Modeling - Dec 2019"
        VersionNumberChanges(39) = "(2.21) see updates described in  - Proposed FRAM Changes for 2021 Pre-season Modeling - Dec 2020"
    End Sub

   Private Sub FVS_Continue_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FVS_Continue.Click
      Me.Hide()
      FVS_MainMenu.ShowDialog()
   End Sub

   Sub StopFormReSize()
      FVS_BackwardsFram_ReSize = True
      FVS_BackwardsResults_ReSize = True
      FVS_BackwardsTarget_ReSize = True
      FVS_BackwardsYearSelect_ReSize = True
      FVS_BasePeriodSelect_ReSize = True
      FVS_EditRecordSetInfo_ReSize = True
      FVS_FisheryMultiGroupSelect_ReSize = True
      FVS_FisheryScalerEdit_ReSize = True
      FVS_FisheryScalerScreen_ReSize = True
      FVS_FisherySelect_ReSize = True
      FVS_FishStkCompScreen_ReSize = True
      FVS_FramUtils_ReSize = True
      FVS_InputMenu_ReSize = True
      FVS_MainMenu_ReSize = True
      FVS_ModelRunSelection_ReSize = True
      FVS_MortalityReport_ReSize = True
      FVS_MortalityTypeSelection_ReSize = True
      FVS_NonRetention_ReSize = True
      FVS_Output_ReSize = True
      FVS_OutputDriver_ReSize = True
      FVS_OutputDriverSelection_ReSize = True
      FVS_PSCCohoERScreen_ReSize = True
      FVS_PSCMaxER_ReSize = True
      FVS_ReportSelection_ReSize = True
      FVS_RunModel_ReSize = True
      FVS_SaveModelRunInputs_ReSize = True
      FVS_ScreenReports_ReSize = True
      FVS_SelectiveFisheryScalerEdit_ReSize = True
      FVS_StockImpactsPer1000Screen_ReSize = True
      FVS_StockRecruitEdit_ReSize = True
      FVS_StockSelect_ReSize = True
      FVS_PopStatScreen_ReSize = True
      FVS_Coweeman_ReSize = True
   End Sub

   Sub StartFormReSize()
      FVS_BackwardsFram_ReSize = False
      FVS_BackwardsResults_ReSize = False
      FVS_BackwardsTarget_ReSize = False
      FVS_BackwardsYearSelect_ReSize = False
      FVS_BasePeriodSelect_ReSize = False
      FVS_EditRecordSetInfo_ReSize = False
      FVS_FisheryMultiGroupSelect_ReSize = False
      FVS_FisheryScalerEdit_ReSize = False
      FVS_FisheryScalerScreen_ReSize = False
      FVS_FisherySelect_ReSize = False
      FVS_FishStkCompScreen_ReSize = False
      FVS_FramUtils_ReSize = False
      FVS_InputMenu_ReSize = False
      FVS_MainMenu_ReSize = False
      FVS_ModelRunSelection_ReSize = False
      FVS_MortalityReport_ReSize = False
      FVS_MortalityTypeSelection_ReSize = False
      FVS_NonRetention_ReSize = False
      FVS_Output_ReSize = False
      FVS_OutputDriver_ReSize = False
      FVS_OutputDriverSelection_ReSize = False
      FVS_PSCCohoERScreen_ReSize = False
      FVS_PSCMaxER_ReSize = False
      FVS_ReportSelection_ReSize = False
      FVS_RunModel_ReSize = False
      FVS_SaveModelRunInputs_ReSize = False
      FVS_ScreenReports_ReSize = False
      FVS_SelectiveFisheryScalerEdit_ReSize = False
      FVS_StockImpactsPer1000Screen_ReSize = False
      FVS_StockRecruitEdit_ReSize = False
      FVS_StockSelect_ReSize = False
      FVS_PopStatScreen_ReSize = False
      FVS_Coweeman_ReSize = False
   End Sub


    Private Sub VersionLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VersionLabel.Click

    End Sub
End Class

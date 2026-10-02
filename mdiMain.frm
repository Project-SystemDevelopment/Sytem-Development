VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   7380
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10875
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   4680
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   7095
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8414
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "06/12/2016"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:12 PM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Page &Setup"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu CloseConnection 
         Caption         =   "&Close Connection"
      End
   End
   Begin VB.Menu Membership 
      Caption         =   "&Membership "
      Index           =   1
      Begin VB.Menu MembershipRegistration 
         Caption         =   "Registratio&n"
      End
      Begin VB.Menu MembershipAdjustment 
         Caption         =   "&Adjustment"
      End
      Begin VB.Menu MembershipEnquiry 
         Caption         =   "&Enquiry"
      End
      Begin VB.Menu MembershipReport 
         Caption         =   "Leaflet && Card"
      End
      Begin VB.Menu frmLeafletBranch 
         Caption         =   "Leaflet Batch"
      End
      Begin VB.Menu mnuCardSTMB 
         Caption         =   "Card Printing (STMB)"
      End
      Begin VB.Menu ZurichCardPrinting 
         Caption         =   "mnuZurichcardPrinting"
      End
      Begin VB.Menu mnuMembershipAXA 
         Caption         =   "Membership List (AXA Affin/Zurich)"
      End
      Begin VB.Menu mnuMBMListETQMBB 
         Caption         =   "Membership List (Etiqa/Maybank)"
      End
      Begin VB.Menu mnucardsend 
         Caption         =   "Card Sending Mode"
      End
      Begin VB.Menu mnuMBMReinstatement 
         Caption         =   "Member Reinstatement"
      End
      Begin VB.Menu MembershipMCOPending 
         Caption         =   "&MCO Pending"
      End
      Begin VB.Menu MembershipDelete 
         Caption         =   "&Deletion"
      End
      Begin VB.Menu MembershipTracking 
         Caption         =   "&Tracking"
      End
      Begin VB.Menu MBMTrackingSimple 
         Caption         =   "Tracking (Simple)"
      End
      Begin VB.Menu MembershipTracking2 
         Caption         =   "Tracking (Jupiter)"
      End
      Begin VB.Menu MBMTrackingSimpleJup 
         Caption         =   "Tracking (Simple - Jupiter)"
      End
      Begin VB.Menu mnuniam 
         Caption         =   "NIAM Report"
      End
      Begin VB.Menu mnuDashm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUplMemberExcel 
         Caption         =   "Upload Membership Excel(Pluto)"
      End
      Begin VB.Menu mnuMemberASCII 
         Caption         =   "Generate Membership ASCII(Pluto)"
      End
      Begin VB.Menu mnuZurichUpload 
         Caption         =   "&Upload Zurich Bordx"
      End
      Begin VB.Menu mnuUpload 
         Caption         =   "&Upload Bordx"
      End
      Begin VB.Menu mnuUpdBordx2013 
         Caption         =   "&Upload Bordx ETQ&&Maybank"
      End
      Begin VB.Menu mnuUploadETQBuss 
         Caption         =   "&Upload Bordx Etiqa Bussiness"
      End
      Begin VB.Menu mnuUpdMBMClinics 
         Caption         =   "&Update Member Panel Clinics"
      End
      Begin VB.Menu mnuMBMAscii 
         Caption         =   "&Generate Ascii"
      End
      Begin VB.Menu mnuGenMAAAscii 
         Caption         =   "Generate MAA ASCII"
      End
      Begin VB.Menu mnDashM 
         Caption         =   "-"
      End
      Begin VB.Menu mnTempMBMReg 
         Caption         =   "Temp MBM Reg"
      End
      Begin VB.Menu mnTempMBMAdj 
         Caption         =   "Temp MBM Adj"
      End
   End
   Begin VB.Menu MCOFees 
      Caption         =   "MCOFees"
      Begin VB.Menu FeeListingNew 
         Caption         =   "MCO Inv/Listing"
      End
      Begin VB.Menu FeeListingNewGST 
         Caption         =   "MCO Inv/Listing (GST)"
      End
      Begin VB.Menu frmManulInvoice 
         Caption         =   "MCO Manual Invoice"
      End
      Begin VB.Menu mnuMCOReinstatement 
         Caption         =   "MCO Reinstatement"
      End
      Begin VB.Menu RefundListing 
         Caption         =   "Refund Listing"
      End
      Begin VB.Menu mnuMCOCorrection 
         Caption         =   "MCO Correction"
      End
      Begin VB.Menu mnuInvoiceDateCorrection 
         Caption         =   "Date Correction"
      End
      Begin VB.Menu InvManualUpd 
         Caption         =   "Invoice &Manual Upd"
      End
      Begin VB.Menu mnuRecMCO 
         Caption         =   "Receipt MCO/Others"
      End
      Begin VB.Menu mnuRecMCOOthMan 
         Caption         =   "Receipt MCO/Others (Manual)"
      End
      Begin VB.Menu mnuForeignWorkers 
         Caption         =   "Foreign Workers"
      End
      Begin VB.Menu mnuMcoInvoiceZurich2015 
         Caption         =   "MCO Invoice (ASCII Zurich)"
      End
      Begin VB.Menu mnuMCOInvASCII 
         Caption         =   "MCO Invoice (ASCII)"
      End
   End
   Begin VB.Menu Admission 
      Caption         =   "&Case"
      Begin VB.Menu mnuRecordDoc 
         Caption         =   "Record Document"
      End
      Begin VB.Menu CaseEntry 
         Caption         =   "Case &Registration"
      End
      Begin VB.Menu CaseAdjustment 
         Caption         =   "Case &Adjustment"
      End
      Begin VB.Menu CASInMBM 
         Caption         =   "Case In&quiry"
      End
      Begin VB.Menu mnCaseTransferTest 
         Caption         =   "Case In&quiry(Test)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnCaseTransfer 
         Caption         =   "Case Trans&fer"
      End
      Begin VB.Menu mnuPanelHospitals 
         Caption         =   "Panel Hospitals"
      End
      Begin VB.Menu mnTempCase 
         Caption         =   "Temp Case Transfer"
      End
      Begin VB.Menu mnuMBMLifetimeLmt 
         Caption         =   "Lifetime Lmt Upgrade"
      End
      Begin VB.Menu CaseAuditReview 
         Caption         =   "Case Audit Review"
      End
      Begin VB.Menu mnuARTracking 
         Caption         =   "Audit Review Tracking"
      End
      Begin VB.Menu CaseTracking 
         Caption         =   "Case &Tracking"
      End
      Begin VB.Menu CaseTracking2 
         Caption         =   "Case Tracking (Jupiter)"
      End
      Begin VB.Menu MMAFees 
         Caption         =   "MMA/CPT"
      End
      Begin VB.Menu mnuPanelHOSPMNI 
         Caption         =   "Panel Hospital - OMNI"
      End
      Begin VB.Menu CasePayClaim 
         Caption         =   "Case Pay&&Claim"
      End
      Begin VB.Menu mnuNonDisclousreInfo 
         Caption         =   "Non-Disclousre Info"
      End
      Begin VB.Menu Correspondence 
         Caption         =   "Correspondence"
      End
      Begin VB.Menu MnuCaseSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCase2Claim 
         Caption         =   "Send to Bordx Unit"
      End
      Begin VB.Menu mnuCase2ClaimRep 
         Caption         =   "Send to Claims (Rep)"
      End
      Begin VB.Menu mnuCRUSend2Claim 
         Caption         =   "CRU Send to Claims"
      End
      Begin VB.Menu mnuLGTracking 
         Caption         =   "&LG Tracking"
      End
      Begin VB.Menu mnuClaimASCII 
         Caption         =   "Clai&m ASCII"
      End
      Begin VB.Menu CaseASCII 
         Caption         =   "Case ASCII"
      End
      Begin VB.Menu mnuClaimASCIIMedGen 
         Caption         =   "Claim ASCII(MedGen)"
      End
      Begin VB.Menu mnuClaimASCIIMedGenNew 
         Caption         =   "Claim ASCII (MedGen - New)"
      End
      Begin VB.Menu mnuClaimASCIIEtiqa 
         Caption         =   "Claims ASCII (Etiqa)"
      End
      Begin VB.Menu mnuClaimASCIIEtiqaSpec 
         Caption         =   "Claims ASCII (Etiqa-Special)"
      End
      Begin VB.Menu mnuAXAAffin 
         Caption         =   "Claim ASCII (AXA Affin)"
      End
      Begin VB.Menu CaseASCIIZur 
         Caption         =   "Case ASCII (Zurich)"
      End
      Begin VB.Menu mnuClaimsASCIIZurichGen 
         Caption         =   "Claims ASCII (Zurich)"
      End
      Begin VB.Menu mnuCLMASCIITIMS 
         Caption         =   "Claims ASCII (TIMS)"
      End
      Begin VB.Menu mnuFrmClaimsAsciiZurich 
         Caption         =   "Claims ASCII Zurich(2015)"
      End
      Begin VB.Menu MnuZurichRejectPending 
         Caption         =   "Claims ASCII Zurich Reject/Pending (2015)"
      End
   End
   Begin VB.Menu MnuInvoices 
      Caption         =   "&Invoice"
      Begin VB.Menu mniInvoiceProcess 
         Caption         =   "Invoice &Processing"
      End
      Begin VB.Menu mnuInvMBMAcc 
         Caption         =   "Invoice &MBM Accounting"
      End
      Begin VB.Menu mnuInvHOSAcc 
         Caption         =   "Invoice &HOS Accounting"
      End
      Begin VB.Menu InvoiceHOSDebittoMBM 
         Caption         =   "Invoice HOS (DR to MBM)"
      End
      Begin VB.Menu InvoiceHOSCRtoMBM 
         Caption         =   "Invoice HOS (CR to MBM)"
      End
   End
   Begin VB.Menu Claim 
      Caption         =   "Work&sheet"
      Begin VB.Menu mnuNewWorksheet 
         Caption         =   "Worksheet &Registration"
      End
      Begin VB.Menu mnuWSAdjust 
         Caption         =   "WorkSheet &Adjustment"
      End
      Begin VB.Menu mnuWSEnq 
         Caption         =   "Worksheet Enquiry"
      End
      Begin VB.Menu mnuCreditDebitNote 
         Caption         =   "Generate Credit/Debit Note"
      End
      Begin VB.Menu mnuWSToAC 
         Caption         =   "Worksheet &Send to A/C"
      End
      Begin VB.Menu mnuWSReturnCLM 
         Caption         =   "Worksheet Send to A/R"
      End
      Begin VB.Menu mnuWksTracking 
         Caption         =   "Worksheet &Tracking"
      End
      Begin VB.Menu mnuWksTracking2 
         Caption         =   "Worksheet Tracking (Jupiter)"
      End
      Begin VB.Menu mnuNewWorksheetGST 
         Caption         =   "Worksheet Registration (GST)"
      End
      Begin VB.Menu mnuWSAdjGST 
         Caption         =   "Worksheet Adjustment (GST)"
      End
   End
   Begin VB.Menu mnuClaim 
      Caption         =   "&Claim"
      Begin VB.Menu ClaimTracking1Jup 
         Caption         =   "Claim Tracking (Sat.)"
         Begin VB.Menu ClaimTracking 
            Caption         =   "Claim Report 1"
         End
         Begin VB.Menu ClaimTracking2 
            Caption         =   "Claim Report 2"
         End
         Begin VB.Menu ClaimTracking3 
            Caption         =   "Claim Report 3"
         End
         Begin VB.Menu ClaimTracking4 
            Caption         =   "Claim Report 4"
         End
         Begin VB.Menu ClaimTrackingSimple 
            Caption         =   "Claim Report (Simple)"
         End
         Begin VB.Menu KPIReport 
            Caption         =   "KPI Report"
         End
      End
      Begin VB.Menu ClaimTracking2Nep 
         Caption         =   "Claim Tracking (Jup.)"
         Begin VB.Menu ClaimTrackingNep 
            Caption         =   "Claim Report 1"
         End
         Begin VB.Menu ClaimTrackingNep2 
            Caption         =   "Claim Report 2"
         End
         Begin VB.Menu ClaimTrackingNep3 
            Caption         =   "Claim Report 3"
         End
         Begin VB.Menu ClaimTrackingNep4 
            Caption         =   "Claim Report 4"
         End
         Begin VB.Menu ClaimTrackingSimpleJup 
            Caption         =   "Claim Report (Simple)"
         End
         Begin VB.Menu KPIReportJup 
            Caption         =   "KPI Report"
         End
      End
      Begin VB.Menu DcuTracking 
         Caption         =   "DCU Tracking"
         Begin VB.Menu mnuReport1 
            Caption         =   "Claim Report"
         End
         Begin VB.Menu mnuReport2 
            Caption         =   "Claim Report (Jup.)"
         End
      End
      Begin VB.Menu mnuClaimNotification 
         Caption         =   "Claim Notification"
      End
      Begin VB.Menu mnuCorrespondence 
         Caption         =   "Correspondence"
      End
      Begin VB.Menu CLMReminder 
         Caption         =   "Claim Reminder"
      End
      Begin VB.Menu mnuClaimPass2AC 
         Caption         =   "Claim Pass to A/C"
      End
      Begin VB.Menu mnuClaimUpload 
         Caption         =   "PMCare Claims Upload"
      End
      Begin VB.Menu mnuEtiqaRpt 
         Caption         =   "Etiqa Report"
      End
   End
   Begin VB.Menu mnuBordx 
      Caption         =   "Claim &Bordx"
      Begin VB.Menu mnuWSSelection 
         Caption         =   "&Worksheet "
         Begin VB.Menu mnuGenerateBordx 
            Caption         =   "&Insert Worksheet"
         End
         Begin VB.Menu mnuBordxRevise 
            Caption         =   "&Remove Worksheet"
         End
      End
      Begin VB.Menu mnuPrintBordx 
         Caption         =   "&Print Bordx"
      End
      Begin VB.Menu mnuPrintRecovery 
         Caption         =   "Print Recovery"
      End
      Begin VB.Menu mnuPaymentAdvice 
         Caption         =   "Payment Advice"
         Begin VB.Menu mnuPayAdviceBordx 
            Caption         =   "Generate Bordx No"
         End
         Begin VB.Menu mnuPayAdviceHosp 
            Caption         =   "Hospital Report"
         End
         Begin VB.Menu mnuGenerateBordxNoAXA 
            Caption         =   "Generate Bordx No - AXA Affin"
         End
         Begin VB.Menu mnuPayAdviceHospAXA 
            Caption         =   "Generate Bordx - AXA Affin"
         End
      End
      Begin VB.Menu mnuPrintClaimInvIBM 
         Caption         =   "Print Claim Invoice(IBM)"
      End
      Begin VB.Menu mnuPrintCLMInvBSN 
         Caption         =   "Print Claim Invoice (BSN)"
      End
      Begin VB.Menu mnuPrintBordxExc 
         Caption         =   "Print Bordx &Excess"
      End
      Begin VB.Menu mnuUploadApproval 
         Caption         =   "Upload Approval/Sent Date"
      End
      Begin VB.Menu UpdateBordx 
         Caption         =   "Update Bordx"
      End
      Begin VB.Menu mnuBordxDelivery 
         Caption         =   "Bor &Delivery "
         Begin VB.Menu mnuBordxSendDate 
            Caption         =   "Bordx Send Date"
         End
         Begin VB.Menu mnuToPayor 
            Caption         =   "Send to &Payor"
            Visible         =   0   'False
         End
         Begin VB.Menu RetfrPay 
            Caption         =   "&Returned From Payor"
         End
         Begin VB.Menu mnuFrPayorCase 
            Caption         =   "Returned Fr Payor(Not App)"
         End
         Begin VB.Menu mnuSentAC 
            Caption         =   "Send to &A/C"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu BorAccPeriod 
         Caption         =   "Bor Acc Period"
      End
      Begin VB.Menu mnsBordxTrack 
         Caption         =   "Moni&toring"
         Begin VB.Menu mnuBordxCtrl 
            Caption         =   "Bordx &Control"
         End
         Begin VB.Menu mnuOSBordx 
            Caption         =   "&Outstanding Bordx"
         End
         Begin VB.Menu mnuBordxTracking 
            Caption         =   "&Bordx && Worksheet"
         End
      End
      Begin VB.Menu mnuApproval 
         Caption         =   "Bordx Approval"
      End
      Begin VB.Menu BordxTracking 
         Caption         =   "Bor &Tracking"
      End
      Begin VB.Menu storage 
         Caption         =   "Storage"
         Begin VB.Menu StorageEntry 
            Caption         =   "Claim files"
         End
         Begin VB.Menu StorageEnquiry 
            Caption         =   "Enquiry"
         End
         Begin VB.Menu StoRepudiateCase 
            Caption         =   "Repudiated Case"
         End
         Begin VB.Menu StoRepEnq 
            Caption         =   "Repudiated Enquiry"
         End
         Begin VB.Menu StoDocument 
            Caption         =   "Document"
         End
         Begin VB.Menu mnuDnRpt 
            Caption         =   "Monthly DN Report"
         End
      End
   End
   Begin VB.Menu mnuPosting 
      Caption         =   "A/C Export"
      Begin VB.Menu InvoiceExp 
         Caption         =   "Invoice"
         Begin VB.Menu Hospital 
            Caption         =   "Hospital"
         End
         Begin VB.Menu Member 
            Caption         =   "Member"
         End
      End
      Begin VB.Menu mnuReceiptExp 
         Caption         =   "Receipt"
         Begin VB.Menu MNRB 
            Caption         =   "MNRB"
         End
         Begin VB.Menu PayorExp 
            Caption         =   "Payor"
         End
         Begin VB.Menu RecExpOthers 
            Caption         =   "Others"
         End
      End
      Begin VB.Menu PostPayment 
         Caption         =   "Payment"
         Begin VB.Menu PostPay2Hosp 
            Caption         =   "Hospital"
         End
         Begin VB.Menu Post2MBM 
            Caption         =   "Member"
         End
      End
      Begin VB.Menu PostBordx 
         Caption         =   "Bordx"
         Begin VB.Menu BordxPost2MNRB 
            Caption         =   "MNRB"
         End
         Begin VB.Menu BordxPost2Pay 
            Caption         =   "Payor"
         End
      End
      Begin VB.Menu mnSales 
         Caption         =   "Sales"
         Begin VB.Menu mnSalesInvoice 
            Caption         =   "Invoice"
         End
      End
      Begin VB.Menu mnuPostingBTBG 
         Caption         =   "BTBG"
      End
   End
   Begin VB.Menu mnuStatistic 
      Caption         =   "&Statistic"
      Visible         =   0   'False
      Begin VB.Menu mnuDateSelection 
         Caption         =   "Date Selection"
      End
   End
   Begin VB.Menu mnuPaymentModule 
      Caption         =   "&Payment"
      Begin VB.Menu mnuPaymentHos 
         Caption         =   "To &Hospital"
      End
      Begin VB.Menu mnuPaymentMBM 
         Caption         =   "To &Member"
         Begin VB.Menu mnPay2MBMAuto 
            Caption         =   "&Pay2MBM(Auto)"
         End
         Begin VB.Menu mnPay2MBM 
            Caption         =   "&Pay2MBM"
         End
         Begin VB.Menu mnPay2MBMReverse 
            Caption         =   "Chq. &Cancellation"
         End
         Begin VB.Menu mnPay2MBMReissue 
            Caption         =   "Chq. &Reissue"
         End
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "Email Notification"
      End
      Begin VB.Menu mnPaymentPayor 
         Caption         =   "&Payor"
      End
      Begin VB.Menu mnPaymentOther 
         Caption         =   "&Others"
      End
      Begin VB.Menu mnuUpdateInvoice 
         Caption         =   "Update &Invoice"
      End
      Begin VB.Menu mnuPayAllocation 
         Caption         =   "PayAllocation"
         Begin VB.Menu mnuMedixGainLoss 
            Caption         =   "&MedixGainLoss"
            Visible         =   0   'False
         End
         Begin VB.Menu mnMedixFloat 
            Caption         =   "Medix &Float"
         End
         Begin VB.Menu mbMedixBankChrg 
            Caption         =   "Medix &Bank Chrg"
         End
         Begin VB.Menu mnMedixGuarantee 
            Caption         =   "Medix &Guarantee"
         End
         Begin VB.Menu mnBankContra 
            Caption         =   "Bank &Contra"
         End
         Begin VB.Menu mnHospAdj 
            Caption         =   "Hosp &Adj"
         End
      End
      Begin VB.Menu mnuPaymentTracking 
         Caption         =   "Payment &Tracking"
      End
      Begin VB.Menu Enquiry 
         Caption         =   "Enquiry"
         Begin VB.Menu PaymentToHospital 
            Caption         =   "Payment To Hospital"
         End
         Begin VB.Menu PaymentToMember 
            Caption         =   "Payment To Member"
         End
      End
      Begin VB.Menu mnuTrustAC 
         Caption         =   "Reconciliation"
         Visible         =   0   'False
      End
      Begin VB.Menu mnPVUpdate 
         Caption         =   "PV Update (Manual)"
      End
      Begin VB.Menu mnuPayAmtAdj 
         Caption         =   "Payable Amt Adjustment"
      End
   End
   Begin VB.Menu mnuReceiptModule 
      Caption         =   "Receipt"
      Begin VB.Menu mnuReceiptMBM 
         Caption         =   "From &Member"
      End
      Begin VB.Menu mnuOthers 
         Caption         =   "&Others"
      End
      Begin VB.Menu mnuReceiptPayor 
         Caption         =   "&Payor"
      End
      Begin VB.Menu mnuReceiptPayorManual 
         Caption         =   "&Payor (Manual)"
      End
      Begin VB.Menu mnuReceipt 
         Caption         =   "&MNRB"
      End
      Begin VB.Menu RecCLMFloatTrans 
         Caption         =   "Claim Float Transfer"
      End
      Begin VB.Menu RECCLMFloatReq 
         Caption         =   "Claim Float Request"
      End
      Begin VB.Menu mnuFloatReceipt 
         Caption         =   "Claim Float Receipt"
      End
      Begin VB.Menu RECCLMFloatTopup 
         Caption         =   "Claim Float Topup"
      End
      Begin VB.Menu mnuAllocation 
         Caption         =   "&RecAllocation"
         Begin VB.Menu mnuRecMedixGainLoss 
            Caption         =   "MedixGainLoss"
         End
         Begin VB.Menu mnuMedixBankCharge 
            Caption         =   "Medix Bank Charge"
         End
         Begin VB.Menu mnuMedixBankContra 
            Caption         =   "Medix Bank Contra"
         End
         Begin VB.Menu mnuMedixHospAdj 
            Caption         =   "Medix Hosp Adj"
         End
      End
      Begin VB.Menu mnuReceiptTracking 
         Caption         =   "Receipt &Tracking"
      End
   End
   Begin VB.Menu mnPosting 
      Caption         =   "Posting"
      Visible         =   0   'False
      Begin VB.Menu mnPostInvoice 
         Caption         =   "Invoice"
         Begin VB.Menu mnPostInvHosp 
            Caption         =   "&Hospital"
         End
         Begin VB.Menu mnPostMBM 
            Caption         =   "&Member"
         End
      End
   End
   Begin VB.Menu MaintenanceTables 
      Caption         =   "Main&tenance"
      Begin VB.Menu MaintenancePlan 
         Caption         =   "&Plan "
      End
      Begin VB.Menu mnuBenefits 
         Caption         =   "&Benefits"
      End
      Begin VB.Menu MaintenanceCase 
         Caption         =   "&Case"
      End
      Begin VB.Menu MnuMemberMaintenance 
         Caption         =   "&Date"
      End
      Begin VB.Menu MaintenanceGeneral 
         Caption         =   "Entity"
      End
      Begin VB.Menu GroupMaintainace 
         Caption         =   "GroupCompany Maintainace"
      End
      Begin VB.Menu MaintenanceOthers 
         Caption         =   "&Member"
      End
      Begin VB.Menu HOSPanelMaintenance 
         Caption         =   "Panel Hospital Maintenance"
      End
      Begin VB.Menu mnuComInfo 
         Caption         =   "Company &Info"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSurgicalOperation 
         Caption         =   "&Surgical Schedule"
      End
      Begin VB.Menu MMA 
         Caption         =   "&MMA"
      End
      Begin VB.Menu mnDataTrans 
         Caption         =   "&Data Transfer"
         Begin VB.Menu mnMBMTransfer 
            Caption         =   "&Member"
         End
         Begin VB.Menu mnCASTrans 
            Caption         =   "&Case"
         End
      End
      Begin VB.Menu SecurityMaintenance 
         Caption         =   "&Security"
         Begin VB.Menu GroupMaintenance 
            Caption         =   "&Department"
         End
         Begin VB.Menu mnuPermission 
            Caption         =   "&User Permission"
         End
      End
   End
   Begin VB.Menu View 
      Caption         =   "&View"
      Visible         =   0   'False
      Begin VB.Menu MnuStatusBar 
         Caption         =   "&Status Bar"
      End
   End
   Begin VB.Menu mnuWindowBar 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Cascade Win&dow"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
'Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_CLOSE = &H10

Private Sub BordxPost2MNRB_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(175) = True Then 'Posting - Bordx - MNRB
        If FormLoadedByName("frmPostingBordxInv2MNRB") = True Then
            frmPostingBordxInv2MNRB.SetFocus
        Else
            frmPostingBordxInv2MNRB.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub BordxPost2Pay_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(176) = True Then 'Posting - Bordx - Payor
        If FormLoadedByName("frmPostingBordxInv2Pay") = True Then
            frmPostingBordxInv2Pay.SetFocus
        Else
            frmPostingBordxInv2Pay.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub CaseASCII_Click()
On Error GoTo errOpen
If CheckAccess(18) = True Then 'Claim ASCII
    If FormLoadedByName("frmCaseASCII") = True Then
        frmCaseASCII.SetFocus
    Else
        frmCaseASCII.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub CaseASCIIZur_Click()
On Error GoTo errOpen
If CheckAccess(18) = True Then 'Claim ASCII
    If FormLoadedByName("frmCaseASCIIZurichGen") = True Then
        frmCaseASCIIZurichGen.SetFocus
    Else
        frmCaseASCIIZurichGen.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub CaseAuditReview_Click()
On Error GoTo errOpen

If CheckAccess(16) = True Then 'Case Audit Review
        If FormLoadedByName("frmCaseAuditReview") = True Then
            frmCaseAuditReview.SetFocus
        Else
            frmCaseAuditReview.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub CasePayClaim_Click()
On Error GoTo errOpen
    
    If CheckAccess(11) = True Then 'Case PayClaim
        If FormLoadedByName("frmCASPayClaim") = True Then
            frmCASPayClaim.SetFocus
        Else
            frmCASPayClaim.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub CaseTracking2_Click()
On Error GoTo errOpen
    If CheckAccess(15) = True Then 'Case Tracking (Neptune)
        If FormLoadedByName("frmCaseTrackingNep") = True Then
            FrmCaseTrackingNep.SetFocus
        Else
            FrmCaseTrackingNep.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub ClaimTracking3_Click()
If CheckTime("TRACKING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(46) = True Then 'Claim Tracking 3
        If FormLoadedByName("FrmClaimRpt3") = True Then
            FrmClaimRpt3.SetFocus
        Else
            FrmClaimRpt3.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub ClaimTracking4_Click()
If CheckTime("TRACKING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(47) = True Then 'Claim Tracking 4
        If FormLoadedByName("FrmClaimRpt4") = True Then
            FrmClaimRpt4.SetFocus
        Else
            FrmClaimRpt4.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub ClaimTrackingNep_Click()
On Error GoTo errOpen
    If CheckAccess(50) = True Then 'Claim Tracking 1 (Neptune)
        If FormLoadedByName("FrmClaimRptNep") = True Then
            FrmClaimRptNep.SetFocus
        Else
            FrmClaimRptNep.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub ClaimTrackingNep2_Click()
On Error GoTo errOpen
    If CheckAccess(51) = True Then 'Claim Tracking 2 (Neptune)
        If FormLoadedByName("FrmClaimRptNep2") = True Then
            FrmClaimRptNep2.SetFocus
        Else
            FrmClaimRptNep2.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub ClaimTrackingNep3_Click()
On Error GoTo errOpen
    If CheckAccess(52) = True Then 'Claim Tracking 3 (Neptune)
        If FormLoadedByName("FrmClaimRptNep3") = True Then
            FrmClaimRptNep3.SetFocus
        Else
            FrmClaimRptNep3.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub ClaimTrackingNep4_Click()
On Error GoTo errOpen
    If CheckAccess(53) = True Then 'Claim Tracking 4 (Neptune)
        If FormLoadedByName("FrmClaimRptNep4") = True Then
            FrmClaimRptNep4.SetFocus
        Else
            FrmClaimRptNep4.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub ClaimTrackingSimple_Click()
    If CheckTime("TRACKING") = False Then Exit Sub
    On Error GoTo errOpen
        If CheckAccess(44) = True Then 'Claim Tracking Simple
            If FormLoadedByName("FrmClaimRptSimple") = True Then
                frmClaimRptSimple.SetFocus
            Else
                frmClaimRptSimple.Show
            End If
        Else
           MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
        Exit Sub
errOpen:
        If Err.Number = 7 Then
            MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
        Else
            MsgBox Err.Description
        End If

End Sub

Private Sub ClaimTrackingSimpleJup_Click()
On Error GoTo errOpen
    If CheckAccess(50) = True Then 'Claim Tracking 1 (Neptune)
        If FormLoadedByName("frmClaimRptSimpleREP") = True Then
            frmClaimRptSimpleREP.SetFocus
        Else
            frmClaimRptSimpleREP.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub CLMReminder_Click()
On Error GoTo errOpen
    If CheckAccess(11) = True Then  'Case Registration
        If FormLoadedByName("frmClaimReminder") = True Then
            frmClaimReminder.SetFocus
        Else
            frmClaimReminder.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub ContraAllocation_Click()
On Error GoTo errOpen
    If CheckAccess(56) = True Then 'Payment to Member
        If FormLoadedByName("frmContraAllocation") = True Then
            frmContraAllocation.SetFocus
        Else
            frmContraAllocation.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub Correspondence_Click()
'On Error GoTo errOpen
'    If CheckAccess(13) = True Then 'Case Inquiry
'        If FormLoadedByName("frmCorrespondence") = True Then
'            frmCorrespondence.SetFocus
'        Else
'            frmCorrespondence.Show
'        End If
'    Else
'       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
'    End If
'    Exit Sub
'errOpen:
'    If Err.Number = 7 Then
'        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
'    Else
'        MsgBox Err.Description
'    End If
End Sub

Private Sub Fax_Click()
On Error GoTo errOpen
    
    If CheckAccess(11) = True Then 'Case Registration
        If FormLoadedByName("frmCaseAdmission") = True Then
            FrmFaxReg.SetFocus
        Else
            FrmFaxReg.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub FeeListingNew_Click()
On Error GoTo errOpen
    If CheckAccess(20) = True Then 'MCO FEE Listing
        If FormLoadedByName("frmMCOFeesListingNew") = True Then
            frmMCOFeesListingNew.SetFocus
        Else
            frmMCOFeesListingNew.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub frmCaseASCII_Click()
On Error GoTo errOpen
If CheckAccess(18) = True Then 'Case ASCII
    If FormLoadedByName("frmCaseASCII") = True Then
        frmCaseASCII.SetFocus
    Else
        frmCaseASCII.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub FeeListingNewGST_Click()
On Error GoTo errOpen
    If CheckAccess(20) = True Then 'MCO FEE Listing
        If FormLoadedByName("frmMCOFeesListingGST") = True Then
            frmMCOFeesListingGST.SetFocus
        Else
            frmMCOFeesListingGST.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub frmLeafletBranch_Click()
On Error GoTo errOpen
    If CheckAccess(4) = True Then 'Leaflet & Card
        If FormLoadedByName("frmPrintMBMBordxBatch") = True Then
            frmPrintMBMBordxBatch.SetFocus
        Else
            frmPrintMBMBordxBatch.Show
        End If
    Else
       MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub frmManulInvoice_Click()
On Error GoTo errOpen
    If CheckAccess(20) = True Then 'MCO FEE Listing
        If FormLoadedByName("FrmManualInvoice") = True Then
            FrmManualInvoice.SetFocus
        Else
            FrmManualInvoice.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub GroupMaintainace_Click()
On Error GoTo errOpen
    
    If CheckAccess(1) = True Then
        If FormLoadedByName("FrmGroupCompany") = True Then
            FrmGroupCompany.SetFocus
        Else
            FrmGroupCompany.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub HOSPanelMaintenance_Click()
On Error GoTo errOpen
    
    If CheckAccess(1) = True Then
        If FormLoadedByName("FrmMaintenancePanelHos") = True Then
            FrmMaintenancePanelHos.SetFocus
        Else
            FrmMaintenancePanelHos.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub Hospital_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(169) = True Then 'Posting - Invoice - Hos
        If FormLoadedByName("frmPostingInvoiceHosp") = True Then
            frmPostingInvoiceHosp.SetFocus
        Else
            frmPostingInvoiceHosp.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub InvManualUpd_Click()
On Error GoTo errOpen
    If CheckAccess(148) = True Then 'MCO Invoice manual Update
        If FormLoadedByName("frmMCOManualUpd") = True Then
            frmMCOManualUpd.SetFocus
        Else
            frmMCOManualUpd.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub InvoiceHOSCRtoMBM_Click()
On Error GoTo errOpen
  If CheckAccess(26) = True Then 'Invoice Accounting
        If FormLoadedByName("frmHOSInvC2MBM") = True Then
            frmHOSInvC2MBM.SetFocus
        Else
            frmHOSInvC2MBM.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
  Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub InvoiceHOSDebittoMBM_Click()
On Error GoTo errOpen
  If CheckAccess(25) = True Then 'Invoice Accounting
        If FormLoadedByName("frmHOSInvD2MBM") = True Then
            frmHOSInvD2MBM.SetFocus
        Else
            frmHOSInvD2MBM.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
  Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub Logout1_Click()
Dim MyTime, MyUser, MyDate
  
    MyUser = Logon
        
    'If Dir("\\NEPTUNE\FileLogin\Login.txt") = "Login.txt" Then
        'Open "\\NEPTUNE\FileLogin\Login.txt" For Append As #1
    If Dir(DocPath & "FileLogin\Login.txt") = "Login.txt" Then
        Open DocPath & "FileLogin\Login.txt" For Append As #1
        Write #1, "Logout";
        Write #1, MyUser;
        Write #1, Date;
        Write #1, Format(Now, "hh:mm:ss am/pm");
        Write #1,
        Close #1
    End If
    Unload Me
    frmLogin.Show

End Sub

Private Sub KPIReport_Click()
'On Error GoTo errOpen
'    If CheckAccess(44) = True Then 'Claim Tracking
'        If FormLoadedByName("frmKPIReportSat") = True Then
'            frmKPIReportSat.SetFocus
'        Else
'            frmKPIReportSat.Show
'        End If
'    Else
'       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
'    End If
'    Exit Sub
'errOpen:
'    If Err.Number = 7 Then
'        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
'    Else
'        MsgBox Err.Description
'    End If
End Sub

Private Sub mbMedixBankChrg_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Then 'Medix Bank Charges
        If FormLoadedByName("FrmAllocMedixBankChrg") = True Then
            FrmAllocMedixBankChrg.SetFocus
        Else
            FrmAllocMedixBankChrg.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub MBMTrackingSimple_Click()
If CheckTime("TRACKING") = False Then Exit Sub
On Error GoTo errOpen
       
    If CheckAccess(6) = True Then 'Membership Tracking Simple
        If FormLoadedByName("frmNewMBMTrackingsimple") = True Then
            frmNewMBMTrackingSimple.SetFocus
        Else
            frmNewMBMTrackingSimple.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub MBMTrackingSimpleJup_Click()


On Error GoTo errOpen
       
    If CheckAccess(7) = True Then 'Membership Tracking (Neptune)
        If FormLoadedByName("frmMBMTrackingRepSimple") = True Then
            frmMBMTrackingRepSimple.SetFocus
        Else
            frmMBMTrackingRepSimple.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub Member_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(170) = True Then 'Posting - Invoice - MBM
        If FormLoadedByName("frmPostingInvoiceMBM") = True Then
            frmPostingInvoiceMBM.SetFocus
        Else
            frmPostingInvoiceMBM.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub MembershipMCOPending_Click()
On Error GoTo errOpen
    If CheckAccess(3) = True Then 'Membership Inquiry
        If FormLoadedByName("frmMBMMCOPending") = True Then
            frmMBMMCOPending.SetFocus
        Else
            frmMBMMCOPending.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub MembershipTracking2_Click()
On Error GoTo errOpen
       
    If CheckAccess(7) = True Then 'Membership Tracking (Neptune)
        If FormLoadedByName("frmMBMTracking2") = True Then
            FrmMBMTracking2.SetFocus
        Else
            FrmMBMTracking2.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub MMA_Click()
On Error GoTo errOpen
    If CheckAccess(147) = True Then 'MMA
        If FormLoadedByName("frmMMA") = True Then
            frmMMA.SetFocus
        Else
            frmMMA.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub MMAFees_Click()
On Error GoTo errOpen
    If CheckAccess(150) = True Then 'MMA Schedule
        If FormLoadedByName("frmMMASchedule") = True Then
            frmMMASchedule.SetFocus
        Else
            frmMMASchedule.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnBankContra_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Then 'Medix Bank Contra
        If FormLoadedByName("FrmAllocBankContra") = True Then
            FrmAllocBankContra.SetFocus
        Else
            FrmAllocBankContra.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub mnCaseTransfer_Click()
On Error GoTo errOpen
    If CheckAccess(11) = True Then 'Case Transfer(CASE)
        If FormLoadedByName("frmCaseTransferMBMNo") = True Then
            frmCaseTransferMBMNo.SetFocus
        Else
            frmCaseTransferMBMNo.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnCaseTransferTest_Click()
On Error GoTo errOpen
    If CheckAccess(13) = True Then 'Case Inquiry
        If FormLoadedByName("frmCaseInquiry2016") = True Then
            frmCaseInquiry2016.SetFocus
        Else
            frmCaseInquiry2016.Show
            If CheckAccess(178) = True Then
                frmCaseInquiry2016.cmdPrintPymt.Enabled = True
            Else
                frmCaseInquiry2016.cmdPrintPymt.Enabled = False
            End If
            
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnCASTrans_Click()
On Error GoTo errOpen
    If CheckAccess(168) = True Then 'Data Transfer(CASE)
        If FormLoadedByName("frmDataTransferCAS1") = True Then
            frmDataTransferCAS1.SetFocus
        Else
            frmDataTransferCAS1.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnClaimDel_Click()
On Error GoTo errOpen
    If CheckAccess(112) = True Then 'Verify password for IT ppl to Access del mbm & del claim
        If FormLoadedByName("FrmITPasswordEntry") = True Then
            FrmITPasswordEntry.SetFocus
        Else
            FrmITPasswordEntry.Show
        End If
    Else
        MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If

Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub mniOriginalInvUpd_Click()
On Error GoTo errOpen
    If CheckAccess(13) = True Then 'Case Inquiry
        If FormLoadedByName("frmClaimUpdateOIR") = True Then
            frmClaimUpdateOIR.SetFocus
        Else
            frmClaimUpdateOIR.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mniPostingDataTracking_Click()

End Sub

Private Sub mnMBMDel_Click()
On Error GoTo errOpen
    If CheckAccess(112) = True Then 'Verify password for IT ppl to Access del mbm & del claim
        If FormLoadedByName("FrmITPasswordEntry") = True Then
            FrmITPasswordEntry.SetFocus
        Else
            FrmITPasswordEntry.Show
        End If
    Else
        MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If

Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub



Private Sub mnMBMTransfer_Click()
On Error GoTo errOpen
    If CheckAccess(168) = True Then 'Data Transfer(MBM)
        If FormLoadedByName("frmDataTransferMBM1") = True Then
            frmDataTransferMBM1.SetFocus
        Else
            frmDataTransferMBM1.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnMedixFloat_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Then 'Medix Float
        If FormLoadedByName("FrmAllocMedixFloat") = True Then
            FrmAllocMedixFloat.SetFocus
        Else
            FrmAllocMedixFloat.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnMedixGuarantee_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Then 'Medix Guarantee
        If FormLoadedByName("FrmAllocMedixGuarantee") = True Then
            FrmAllocMedixGuarantee.SetFocus
        Else
            FrmAllocMedixGuarantee.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub mnPay2MBM_Click()
    On Error GoTo errOpen
    If CheckAccess(56) = True Then 'Payment to Member
        If FormLoadedByName("frmPayMember") = True Then
            FrmPayMember.SetFocus
        Else
            FrmPayMember.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnPay2MBMAuto_Click()
    On Error GoTo errOpen
    If CheckAccess(56) = True Then 'Payment to Member
        If FormLoadedByName("frmPayMemberJul08") = True Then
            frmPayMemberJul08.SetFocus
        Else
            frmPayMemberJul08.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnPay2MBMReissue_Click()
On Error GoTo errOpen
    If CheckAccess(56) = True Then 'Payment to Member - Reissue
        If FormLoadedByName("frmPaymentMBMReissue") = True Then
            frmPaymentMBMReissue.SetFocus
        Else
            frmPaymentMBMReissue.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnPay2MBMReverse_Click()
On Error GoTo errOpen
    If CheckAccess(56) = True Then 'Payment to Member - cancellation
        If FormLoadedByName("FrmPaymentMBMReverse") = True Then
            FrmPaymentMBMReverse.SetFocus
        Else
            FrmPaymentMBMReverse.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnPaymentOther_Click()

On Error GoTo errOpen
    If CheckAccess(56) = True Then 'Payment to Member
        If FormLoadedByName("FrmPaymentOthers") = True Then
            FrmPaymentOthers.SetFocus
        Else
            FrmPaymentOthers.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub



Private Sub mnPaymentPayor_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Then 'Payment to Payor
        If FormLoadedByName("FrmPaymentPayor") = True Then
            FrmPaymentPayor.SetFocus
        Else
            FrmPaymentPayor.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnPostInvHosp_Click()
On Error GoTo errOpen
If CheckAccess(55) = True Then 'Hosp-Invoice-Posting
    If FormLoadedByName("frmPostingInvoiceHosp") = True Then
        frmPostingInvoiceHosp.SetFocus
    Else
        frmPostingInvoiceHosp.Show
    End If
Else
    MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
End If
Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnPostMBM_Click()
On Error GoTo errOpen
If CheckAccess(56) = True Then 'MBM-Invoice-Posting
    If FormLoadedByName("frmPostingInvoiceMBM") = True Then
        frmPostingInvoiceMBM.SetFocus
    Else
        frmPostingInvoiceMBM.Show
    End If
Else
    MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
End If
Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnPVUpdate_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Or CheckAccess(56) Then 'Payment to Hospital
        If FormLoadedByName("FrmPaymentManual") = True Then
            FrmPaymentManual.SetFocus
        Else
            FrmPaymentManual.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub MNRB_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(171) = True Then 'Posting - Receipt - MNRB
        If FormLoadedByName("frmPostingRecMNRB") = True Then
            frmPostingRecMNRB.SetFocus
        Else
            frmPostingRecMNRB.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnSalesInvoice_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(177) = True Then 'Posting - Sales - Invoice
        If FormLoadedByName("frmPostingSalesGST") = True Then
            'frmPostingSalesInvoice.SetFocus
            frmPostingSalesGST.SetFocus
        Else
            'frmPostingSalesInvoice.Show
            frmPostingSalesGST.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnTempCase_Click()
On Error GoTo errOpen
    
    If CheckAccess(11) = True Then 'Temp Case transfer
        If FormLoadedByName("FrmTempCaseTransfer") = True Then
            FrmTempCaseTransfer.SetFocus
        Else
            FrmTempCaseTransfer.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnTempMBMAdj_Click()
    If FormLoadedByName("FrmTempMBMAdj") = True Then
        FrmTempMBMAdj.SetFocus
    Else
        FrmTempMBMAdj.Show
    End If
End Sub

Private Sub mnTempMBMReg_Click()
    If FormLoadedByName("FrmTempMBMReg") = True Then
        FrmTempMBMReg.SetFocus
    Else
        FrmTempMBMReg.Show
    End If
End Sub

Private Sub mnuApproval_Click()
On Error GoTo errOpen
    If CheckAccess(34) = True Then 'Bordx Approval
        If FormLoadedByName("FrmBordxApproval") = True Then
            FrmBordxApproval.SetFocus
        Else
            FrmBordxApproval.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuARTracking_Click()
On Error GoTo errOpen

If CheckAccess(16) = True Then 'Audit Review Tracking
        If FormLoadedByName("FrmAuditReviewReport") = True Then
            FrmAuditReviewReport.SetFocus
        Else
            FrmAuditReviewReport.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuAttendance_Click()
On Error GoTo errOpen
    If CheckAccess(3) = True Then
        If FormLoadedByName("frmStaffAttendance") = True Then
            frmStaffAttendance.SetFocus
        Else
            frmStaffAttendance.Show vbModal
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuAXAAffin_Click()
On Error GoTo errOpen
If CheckAccess(18) = True Then 'Claim ASCII
    If FormLoadedByName("frmClaimsUploadAXA") = True Then
        frmClaimsUploadAXA.SetFocus
    Else
        frmClaimsUploadAXA.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuBordxSendDate_Click()
On Error GoTo errOpen
    If CheckAccess(35) = True Then 'Set Delivery Date (Send to Payor)
        If FormLoadedByName("frmSendClaimBordxv2") = True Then
            frmSendClaimBordxV2.SetFocus
        Else
            frmSendClaimBordxV2.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnucardsend_Click()
On Error GoTo errOpen
    If CheckAccess(1) = True Then 'Membership Registration
        If FormLoadedByName("frmMembershipCardControl") = True Then
            frmMembershipCardControl.SetFocus
        Else
            frmMembershipCardControl.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuCardSTMB_Click()
On Error GoTo errOpen
    If CheckAccess(4) = True Then 'Leaflet & Card
        If FormLoadedByName("frmMBMCardExpSTMB") = True Then
            frmMBMCardExpSTMB.SetFocus
        Else
            frmMBMCardExpSTMB.Show
        End If
    Else
       MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuCase2Claim_Click()
On Error GoTo errOpen
    If CheckAccess(154) = True Then 'Cases Send to Claims
        If FormLoadedByName("FrmCSSendtoClaim") = True Then
            FrmCSSendtoClaim.SetFocus
        Else
            FrmCSSendtoClaim.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuCase2ClaimRep_Click()
On Error GoTo errOpen
    If CheckAccess(154) = True Then 'Cases Send to Claims
        If FormLoadedByName("FrmCSRepSendtoClaim") = True Then
            FrmCSRepSendtoClaim.SetFocus
        Else
            FrmCSRepSendtoClaim.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuClaimASCIIEtiqa_Click()
On Error GoTo errOpen
If CheckAccess(18) = True Then 'Claim ASCII
    If FormLoadedByName("frmCLMASCIIEtiqa") = True Then
        frmCLMASCIIEtiqa.SetFocus
    Else
        frmCLMASCIIEtiqa.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuClaimASCIIEtiqaSpec_Click()
On Error GoTo errOpen
If CheckAccess(18) = True Then 'Claim ASCII
    If FormLoadedByName("frmCLMASCIIEtiqaSpec") = True Then
        frmCLMASCIIEtiqaSpec.SetFocus
    Else
        frmCLMASCIIEtiqaSpec.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuClaimASCIIMedGen_Click()
On Error GoTo errOpen
    If CheckAccess(18) = True Then 'Claim ASCII
        If FormLoadedByName("frmClaimASCIIMedGen") = True Then
            frmClaimASCIIMedGen.SetFocus
        Else
            frmClaimASCIIMedGen.Show
        End If
    Else
        MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuClaimASCIIMedGenNew_Click()
On Error GoTo errOpen
    If CheckAccess(18) = True Then 'Claim ASCII
        If FormLoadedByName("frmClaimASCIIMedGenNew") = True Then
            frmClaimASCIIMedGenNew.SetFocus
        Else
            frmClaimASCIIMedGenNew.Show
        End If
    Else
        MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuClaimNotification_Click()
On Error GoTo errOpen
    If CheckAccess(11) = True Then 'Case Registration
        If FormLoadedByName("frmClaimNotification") = True Then
            frmClaimNotification.SetFocus
        Else
            frmClaimNotification.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuClaimsASCIIZurichGen_Click()

On Error GoTo errOpen
If CheckAccess(18) = True Then 'Claim ASCII
    If FormLoadedByName("frmClaimsASCIIZurichGen") = True Then
        frmClaimsASCIIZurichGen.SetFocus
    Else
        frmClaimsASCIIZurichGen.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuClaimUpload_Click()
If CheckAccess(64) = True Then 'Claim ASCII
    If FormLoadedByName("FrmPMCareClaimsUpload") = True Then
        FrmPMCareClaimsUpload.SetFocus
    Else
        FrmPMCareClaimsUpload.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuCLMASCIITIMS_Click()

On Error GoTo errOpen
If CheckAccess(12) = True Then 'Claim ASCII
    If FormLoadedByName("frmASCIIClaimTIMS") = True Then
        frmASCIIClaimTIMS.SetFocus
    Else
        frmASCIIClaimTIMS.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuCorrespondence_Click()
On Error GoTo errOpen
    If CheckAccess(11) = True Then 'Case Registration
        If FormLoadedByName("frmCLMCrpd") = True Then
            frmCLMCrpd.SetFocus
        Else
            frmCLMCrpd.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuCreditDebitNote_Click()
On Error GoTo errOpen
    If CheckAccess(65) = True Then 'Worksheet Adjustment
        If FormLoadedByName("frmGenerateCNNote") = True Then
            frmGenerateCNNote.SetFocus
        Else
            frmGenerateCNNote.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuCRUSend2Claim_Click()
On Error GoTo errOpen
    If CheckAccess(154) = True Then 'CRU Send to Claims
        If FormLoadedByName("FrmCRUPassOut") = True Then
            FrmCRUPassOut.SetFocus
        Else
            FrmCRUPassOut.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuCustomerService_Click()

End Sub

Private Sub mnuDnRpt_Click()
On Error GoTo errOpen

If CheckAccess(16) = True Then 'Monthly DN Report
        If FormLoadedByName("FrmMonthlyDNRpt") = True Then
            FrmMonthlyDNRpt.SetFocus
        Else
            FrmMonthlyDNRpt.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub mnuEmail_Click()
    On Error GoTo errOpen
    If CheckAccess(56) = True Then 'Payment to Member
        If FormLoadedByName("frmEmailNotification") = True Then
            frmEmailNotification.SetFocus
        Else
            frmEmailNotification.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuEtiqaRpt_Click()
On Error GoTo errOpen
    If CheckAccess(65) = True Then 'Worksheet Adjustment
        If FormLoadedByName("frmGenerateEtiqaCLM") = True Then
            frmGenerateEtiqaCLM.SetFocus
        Else
            frmGenerateEtiqaCLM.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub





Private Sub mnuFloatReceipt_Click()
On Error GoTo errOpen
    If CheckAccess(59) = True Then 'Receipt from Member
        If FormLoadedByName("frmCLMFloatReceipt") = True Then
            frmCLMFloatReceipt.SetFocus
        Else
            frmCLMFloatReceipt.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuFrmClaimsAsciiZurich_Click()
On Error GoTo errOpen
    If CheckAccess(18) = True Then 'Claim ASCII
        If FormLoadedByName("frmClaimASCIIZurich2015") = True Then
            frmClaimASCIIZurich2015.SetFocus
        Else
            frmClaimASCIIZurich2015.Show
        End If
    Else
        MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuFrPayorCase_Click()
On Error GoTo errOpen
    If CheckAccess(36) = True Then 'Set Delivery Date (Returned From Payor By Claims)
        If FormLoadedByName("FrmReturnPayorCase") = True Then
            FrmReturnPayorCase.SetFocus
        Else
            FrmReturnPayorCase.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuGenerateBordxNoAXA_Click()
On Error GoTo errOpen
    If CheckAccess(30) = True Then 'Insert Worksheet
        If FormLoadedByName("frmPaymentAdviceBordxAXA") = True Then
            frmPaymentAdviceBordxAXA.SetFocus
        Else
            frmPaymentAdviceBordxAXA.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuGenMAAAscii_Click()
On Error GoTo errOpen
    If CheckAccess(9) = True Then 'Generate MBM ASCII
        If FormLoadedByName("frmMAAASCII") = True Then
            frmMAAASCII.SetFocus
        Else
            frmMAAASCII.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
 Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuInvoiceDateCorrection_Click()
On Error GoTo errOpen
    If CheckAccess(20) = True Then 'MCO FEE Listing
        If FormLoadedByName("frmMCOInvoiceCorrection") = True Then
            frmMCOInvoiceCorrection.SetFocus
        Else
            frmMCOInvoiceCorrection.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub



Private Sub mnuInvTrackNep_Click()
On Error GoTo errOpen
  If CheckAccess(28) = True Then 'Invoice Tracking (Neptune)
        If FormLoadedByName("frmInvoiceTrackNep") = True Then
            FrmInvoiceTrackNep.SetFocus
        Else
            FrmInvoiceTrackNep.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
  Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuITActivities_Click()
On Error GoTo errOpen
    If CheckAccess(112) = True Then 'Verify password for IT ppl to Access del mbm & del claim
        If FormLoadedByName("FrmITPasswordEntry") = True Then
            FrmITPasswordEntry.SetFocus
        Else
            FrmITPasswordEntry.Show
        End If
    Else
        MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If

Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuLGDel_Click()
On Error GoTo errOpen
    If CheckAccess(112) = True Then 'Verify password for IT ppl to Access del mbm & del claim
        If FormLoadedByName("FrmITPasswordEntry") = True Then
            FrmITPasswordEntry.SetFocus
        Else
            FrmITPasswordEntry.Show
        End If
    Else
        MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If

Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuMBMLifetimeLmt_Click()
On Error GoTo errOpen
    
    If CheckAccess(11) = True Then 'Case Registration
        If FormLoadedByName("frmMembershipLifetimeLmt") = True Then
            frmMembershipLifetimeLmt.SetFocus
        Else
            frmMembershipLifetimeLmt.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub mnuMBMReinstatement_Click()
On Error GoTo errOpen
    If CheckAccess(1) = True Then 'Membership Registration
        If FormLoadedByName("frmMBMReintstatement") = True Then
            frmMBMReintstatement.SetFocus
        Else
            frmMBMReintstatement.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuMCOCorrection_Click()
On Error GoTo errOpen
    If CheckAccess(20) = True Then 'MCO FEE Listing
        If FormLoadedByName("frmMCOCorrection") = True Then
            frmMCOCorrection.SetFocus
        Else
            frmMCOCorrection.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuMCOInvASCII_Click()
On Error GoTo errOpen
    If CheckAccess(20) = True Then 'MCO FEE Listing
        If FormLoadedByName("frmMCOInvoiceASCII") = True Then
            frmMCOInvoiceASCII.SetFocus
        Else
            frmMCOInvoiceASCII.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuMcoInvoiceZurich2015_Click()
On Error GoTo errOpen
    If CheckAccess(20) = True Then 'MCO FEE Listing
        If FormLoadedByName("frmMCOInvoiceZurich2015") = True Then
            frmMCOInvoiceZurich2015.SetFocus
        Else
            frmMCOInvoiceZurich2015.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuMCOReinstatement_Click()
On Error GoTo errOpen
    If CheckAccess(1) = True Then 'Membership Registration
        If FormLoadedByName("frmMCOInvoiceReinstatement") = True Then
            frmMCOInvoiceReinstatement.SetFocus
        Else
            frmMCOInvoiceReinstatement.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuMedixBankCharge_Click()
'On Error GoTo errOpen
    
    'If CheckAccess(63) = True Then 'Medix Bank Charge
        'If FormLoadedByName("FrmRecBankCharges") = True Then
            'FrmRecBankCharges.SetFocus
        'Else
            FrmRecBankCharges.Show
        'End If
    'Else
       'MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    'End If
    'Exit Sub
'errOpen:
    'If ERR.Number = 7 Then
        'MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    'Else
        'MsgBox ERR.Description
    'End If
End Sub

Private Sub mnuMedixBankContra_Click()
'On Error GoTo errOpen
    
    'If CheckAccess(63) = True Then 'Medix Bank Contra
        'If FormLoadedByName("FrmRecBankContra") = True Then
            'FrmRecBankContra.SetFocus
        'Else
            FrmRecBankContra.Show
        'End If
    'Else
       'MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    'End If
    'Exit Sub
'errOpen:
    'If ERR.Number = 7 Then
        'MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    'Else
        'MsgBox ERR.Description
    'End If
End Sub

Private Sub mnuMedixGainLoss_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Then 'Medix Gain Loss
        If FormLoadedByName("frmAllocMedixGainLoss") = True Then
            frmAllocMedixGainLoss.SetFocus
        Else
            frmAllocMedixGainLoss.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuMedixHospAdj_Click()
'On Error GoTo errOpen
    
    'If CheckAccess(63) = True Then 'Medix Hosp Adj
        'If FormLoadedByName("FrmRecHospAdj") = True Then
            'FrmRecHospAdj.SetFocus
        'Else
            FrmRecHospAdj.Show
        'End If
    'Else
       'MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    'End If
    'Exit Sub
'errOpen:
    'If ERR.Number = 7 Then
        'MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    'Else
        'MsgBox ERR.Description
    'End If
End Sub


Private Sub mnuMemberASCII_Click()
On Error GoTo errOpen
    If CheckAccess(1) = True Then 'Membership Registration
        If FormLoadedByName("frmMembershipASCIIPluto") = True Then
            frmMembershipASCIIPluto.SetFocus
        Else
            frmMembershipASCIIPluto.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuMembershipAXA_Click()
On Error GoTo errOpen
    
    If CheckAccess(2) = True Then 'Membership Adjustment
        If FormLoadedByName("frmMembershipExportAXA") = True Then
            frmMembershipExportAXA.SetFocus
        Else
            frmMembershipExportAXA.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuNewWorksheetGST_Click()
On Error GoTo errOpen
    If CheckAccess(64) = True Then 'Worksheet Registration
        If FormLoadedByName("frmWSRegGST2015V2") = True Then
            frmWSRegGST2015V2.SetFocus
        Else
            frmWSRegGST2015V2.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuNiam_Click()
On Error GoTo errOpen
    If CheckAccess(149) = True Then 'NIAM Report
        If FormLoadedByName("frmNiamReport") = True Then
            FrmNiamReport.SetFocus
        Else
            FrmNiamReport.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub mnuNonDisclousreInfo_Click()
On Error GoTo errOpen
    
    If CheckAccess(11) = True Then 'Case PayClaim
        If FormLoadedByName("frmRECNonDisclosure") = True Then
            frmRECNonDisclosure.SetFocus
        Else
            frmRECNonDisclosure.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuOthers_Click()
On Error GoTo errOpen
    If CheckAccess(60) = True Then 'Receipt from Others
        If FormLoadedByName("frmReceiptOther") = True Then
            FrmReceiptOther.SetFocus
        Else
            FrmReceiptOther.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuPageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .Flags = cdlPDNoSelection + cdlPDNoPageNums
        .ShowPrinter
    End With
End Sub

Private Sub mnuPanelHospitals_Click()
On Error GoTo errOpen
    If CheckAccess(3) = True Then 'Panel Hospitals
        If FormLoadedByName("frmPanelHospital") = True Then
            frmPanelHospital.SetFocus
        Else
            frmPanelHospital.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuPanelHOSPMNI_Click()
On Error GoTo errOpen
    If CheckAccess(13) = True Then 'Case Inquiry
        If FormLoadedByName("frmHOSASCIIFileOMNI") = True Then
            frmHOSASCIIFileOMNI.SetFocus
        Else
            frmHOSASCIIFileOMNI.Show
            'If CheckAccess(178) = True Then
            '    frmHOSASCIIFileOMNI.cmdPrintPymt.Enabled = True
            'Else
            '    frmHOSASCIIFileOMNI.cmdPrintPymt.Enabled = False
            'End If
            
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuPayAdviceBordx_Click()
On Error GoTo errOpen
    If CheckAccess(30) = True Then 'Insert Worksheet
        If FormLoadedByName("frmPayAdviceHosp") = True Then
            frmPayAdviceHosp.SetFocus
        Else
            frmPayAdviceHosp.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuPayAdviceHosp_Click()
On Error GoTo errOpen
    If CheckAccess(30) = True Then 'Insert Worksheet
        If FormLoadedByName("frmPayAdviceHosRpt") = True Then
            frmPayAdviceHosRpt.SetFocus
        Else
            frmPayAdviceHosRpt.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuPayAdviceHospAXA_Click()
On Error GoTo errOpen
    If CheckAccess(30) = True Then 'Insert Worksheet
        If FormLoadedByName("frmPayAdviceHosRptAXA") = True Then
            frmPayAdviceHosRptAXA.SetFocus
        Else
            frmPayAdviceHosRptAXA.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuPayAmtAdj_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Then
        If FormLoadedByName("frmPayableAmtAdj") = True Then
            frmPayableAmtAdj.SetFocus
        Else
            frmPayableAmtAdj.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnuPostingBTBG_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(170) = True Then 'Posting - Invoice - MBM
        If FormLoadedByName("frmPostingBTBG") = True Then
            frmPostingBTBG.SetFocus
        Else
            frmPostingBTBG.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuPrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub

    With dlgCommonDialog
        .DialogTitle = "Print Form"
        .CancelError = True
        .Flags = cdlPDNoSelection + cdlPDNoPageNums
        '.Flags = cdlPDSelection + cdlPDNoPageNums
        
        .ShowPrinter
        
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.PrintForm
        End If
    End With

End Sub


Private Sub mnuPrintBordxExc_Click()
On Error GoTo errOpen
    If CheckAccess(33) = True Then 'Print Bordx Excess
        If FormLoadedByName("frmPrintBordxEx") = True Then
            FrmPrintBordxEx.SetFocus
        Else
            FrmPrintBordxEx.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub mnuPrintClaimInvIBM_Click()
On Error GoTo errOpen
    If CheckAccess(32) = True Then 'Print Bordx
        If FormLoadedByName("frmPrintBordxIBM") = True Then
            frmPrintBordxIBM.SetFocus
        Else
            frmPrintBordxIBM.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub



Private Sub mnuPrintCLMInvBSN_Click()
On Error GoTo errOpen
    If CheckAccess(32) = True Then 'Print Bordx
        If FormLoadedByName("frmPrintBordxBSN") = True Then
            frmPrintBordxBSN.SetFocus
        Else
            frmPrintBordxBSN.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuPrintRecovery_Click()
On Error GoTo errOpen
    If CheckAccess(32) = True Then 'Print Bordx
        If FormLoadedByName("frmGenerateRecovery") = True Then
            frmGenerateRecovery.SetFocus
        Else
            frmGenerateRecovery.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuReceiptManual_Click()
On Error GoTo errOpen
    If CheckAccess(62) = True Then 'Receipt from MNRB - Manual
        If FormLoadedByName("FrmReceiptMNRBManual") = True Then
            FrmReceiptMNRBManual.SetFocus
        Else
            FrmReceiptMNRBManual.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuReceiptPayor_Click()
On Error GoTo errOpen
    If CheckAccess(61) = True Then 'Receipt from Payor
        If FormLoadedByName("frmReceiptPayor") = True Then
            FrmReceiptPayor.SetFocus
        Else
            FrmReceiptPayor.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuReceiptPayorManual_Click()
On Error GoTo errOpen
    If CheckAccess(61) = True Then 'Receipt from Payor (Manual)
        If FormLoadedByName("frmReceiptPayorManual") = True Then
            FrmReceiptPayorManual.SetFocus
        Else
            FrmReceiptPayorManual.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuRecMCO_Click()
On Error GoTo errOpen
    If CheckAccess(19) = True Then 'Receipt MCO/Others
        If FormLoadedByName("frmRecMCO") = True Then
            FrmRecMCO.SetFocus
        Else
            FrmRecMCO.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuRecMCOOthMan_Click()
On Error GoTo errOpen
    If CheckAccess(19) = True Then 'Receipt MCO/Others Manual
        If FormLoadedByName("frmRecMCOManual") = True Then
            frmRecMCOManual.SetFocus
        Else
            frmRecMCOManual.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuRecMedixGainLoss_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Then 'Medix Gain Loss
        If FormLoadedByName("frmRecAllocMedixGainLoss") = True Then
            frmRecAllocMedixGainLoss.SetFocus
        Else
            frmRecAllocMedixGainLoss.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuRecordDoc_Click()
On Error GoTo errOpen
    
    If CheckAccess(11) = True Then 'Case Registration
        If FormLoadedByName("frmCaseRecording") = True Then
            frmCaseRecording.SetFocus
        Else
            frmCaseRecording.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuReport1_Click()
On Error GoTo errOpen
    If CheckAccess(155) = True Then 'DCU Report 1
        If FormLoadedByName("FrmDCUReport1") = True Then
            FrmDCUReport1.SetFocus
        Else
            FrmDCUReport1.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuReport2_Click()
On Error GoTo errOpen
    If CheckAccess(156) = True Then 'DCU Report 2
        If FormLoadedByName("FrmDCUReport2") = True Then
            FrmDCUReport2.SetFocus
        Else
            FrmDCUReport2.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuSunHISPVCheck_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Then
        If FormLoadedByName("frmPVSunHISFlag") = True Then
            frmPVSunHISFlag.SetFocus
        Else
            frmPVSunHISFlag.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuSurgicalOperation_Click()
On Error GoTo errOpen
    If CheckAccess(150) = True Then 'MMA Schedule
        If FormLoadedByName("frmMMASchedule") = True Then
            frmMMASchedule.SetFocus
        Else
            frmMMASchedule.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnuTrustAC_Click()
On Error GoTo errOpen
    If CheckAccess(157) = True Then 'Trust Account
        If FormLoadedByName("FrmTrustAC") = True Then
            FrmTrustAC.SetFocus
        Else
            FrmTrustAC.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuUpdateInvoice_Click()
On Error GoTo errOpen
    If CheckAccess(57) = True Then 'Payment to Hospital
        If FormLoadedByName("frmUpdateInvoice") = True Then
            frmUpdateInvoice.SetFocus
        Else
            frmUpdateInvoice.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub
Private Sub mnuUpdBordx2013_Click()
On Error GoTo errOpen
    If CheckAccess(8) = True Then 'Upload Bordx
        If FormLoadedByName("frmMBMDataUploadEtiqa.show") = True Then
             frmMBMDataUploadEtiqa.SetFocus
        Else
             frmMBMDataUploadEtiqa.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
 Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuUpdMBMClinics_Click()
On Error GoTo errOpen
    If CheckAccess(3) = True Then 'Membership Inquiry
        If FormLoadedByName("frmMemberClinics") = True Then
            frmMemberClinics.SetFocus
        Else
            frmMemberClinics.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuUplMemberExcel_Click()
On Error GoTo errOpen
    If CheckAccess(1) = True Then 'Membership Registration
        If FormLoadedByName("frmMembershipExcelPluto") = True Then
            frmMembershipExcelPluto.SetFocus
        Else
            frmMembershipExcelPluto.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuUploadApproval_Click()
On Error GoTo errOpen
    If CheckAccess(32) = True Then 'Print Bordx
        If FormLoadedByName("frmUploadApproval") = True Then
            frmUploadApproval.SetFocus
        Else
            frmUploadApproval.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuUploadBordx08_Click()

End Sub

Private Sub mnuUploadBordx2010_Click()
On Error GoTo errOpen
    If CheckAccess(8) = True Then 'Upload Bordx
        If FormLoadedByName("FrmNewMBMEndt10.show") = True Then
'            frmNewEndorsement.SetFocus
'            FrmNewEndt.SetFocus
'             FrmNewMBMEndt.SetFocus
             FrmNewMBMEndt10.SetFocus
        Else
'            frmNewEndorsement.show
'            FrmNewEndt.show
'             FrmNewMBMEndt.show
             FrmNewMBMEndt10.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
 Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuUploadETQBuss_Click()
On Error GoTo errOpen
    If CheckAccess(8) = True Then 'Upload Bordx
        If FormLoadedByName("frmMBMDataUploadETQBuss.show") = True Then
'            frmNewEndorsement.SetFocus
'            FrmNewEndt.SetFocus
'             FrmNewMBMEndt.SetFocus
             frmMBMDataUploadETQBuss.SetFocus
        Else
'            frmNewEndorsement.show
'            FrmNewEndt.show
'             FrmNewMBMEndt.show
             frmMBMDataUploadETQBuss.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
 Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub MnuStatusBar_Click()
    MnuStatusBar.Checked = Not MnuStatusBar.Checked
    StatusBar1.Visible = MnuStatusBar.Checked
End Sub

Private Sub MDIForm_Load()
        
    Dim hSysMenu As Long

    hSysMenu = GetSystemMenu(hwnd, False)
    RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
    Me.Caption = Version
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.Caption = DialogBoxTitle & " - " & "Main Menu"
    'StatusBar1.Panels(1).Text = Version
    'StatusBar1.Panels(5).Text = "Logon : " & Logon
    'StatusBar1.Visible = False
    'Me.Arrange vbCascade
End Sub

Private Sub Logout_Click()
Dim MyTime, MyUser, MyDate
  
    MyUser = Logon
        
    If Dir(DocPath & "FileLogin\Login.txt") = "Login.txt" Then
        Open DocPath & "FileLogin\Login.txt" For Append As #1
        Write #1, "Logout";
        Write #1, MyUser;
        Write #1, Date;
        Write #1, Format(Now, "hh:mm:ss am/pm");
        Write #1,
        Close #1
    End If
    
    'For Each Form In Forms
    '    Unload Form
    '    Set Form = Nothing
    'Next Form
    
    'Unload Me
    
    frmLogin.Show


End Sub

Private Sub CloseConnection_Click()
On Error Resume Next
Dim MyTime, MyUser, MyDate
Dim i As Integer
Dim Form As Form
Dim WinWnd As Long
    
    
    
    MyUser = Logon
        
    medixmasterconnA.Close
    medixmasterconn.Close
    medixmasterconnWS.Close
    medixmasterconnPayment.Close
    medixmasterconnMaint.Close
    MediConnCISDB.Close
        
    Set medixmasterconnA = Nothing
    Set medixmasterconn = Nothing
    Set medixmasterconnWS = Nothing
    Set medixmasterconnPayment = Nothing
    Set medixmasterconnMaint = Nothing
    Set MediConnCISDB = Nothing
        
    'If Dir("\\NEPTUNE\FileLogin\Login.txt") = "Login.txt" Then
        'Open "\\NEPTUNE\FileLogin\Login.txt" For Append As #1
    If Dir(DocPath & "FileLogin\Login.txt") = "Login.txt" Then
        Open DocPath & "FileLogin\Login.txt" For Append As #1
        Write #1, "Logout";
        Write #1, MyUser;
        Write #1, Date;
        Write #1, Format(Now, "hh:mm:ss am/pm");
        Write #1,
        Close #1
    End If


   ' Loop through the forms collection and unload
   ' each form.
   
   For i = (Forms.Count) To 0 Step -1
      Unload Forms(i)
      Set Forms(i) = Nothing
   Next
   
   
    
   ' For Each Form In Forms
   '     Unload Form
   '     Set Form = Nothing
   ' Next Form

    
    medixmasterconn.Close
    medixmasterconnA.Close
    medixmasterconnWS.Close
    medixmasterconnPayment.Close
    medixmasterconnMaint.Close
   

    'WinWnd = FindWindow(vbNullString, "medixhis")

    'If WinWnd <> 0 Then
    '    PostMessage WinWnd, WM_CLOSE, 0&, 0&
    ''Else
    ''    MsgBox "No window of that name exists."
    'End If
   
End Sub

' ================= Claim Bordx ==================================

Private Sub mnuOSBordx_Click()
On Error GoTo errOpen
    If CheckAccess(41) = True Then 'Outstanding Bordx
        If FormLoadedByName("frmOSBordx") = True Then
            FrmOSBordx.SetFocus
        Else
            FrmOSBordx.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
 Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
       
End Sub

Private Sub mnuSentAC_Click()
On Error GoTo errOpen
    If CheckAccess(38) = True Then 'Sent to A/C
        If FormLoadedByName("frmSendtoAC") = True Then
            FrmSendtoAC.SetFocus
        Else
            FrmSendtoAC.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuToMNRB_Click()
On Error GoTo errOpen
    If CheckAccess(37) = True Then 'Set Delivery Date (Send to MNRB)
        If FormLoadedByName("frmSendtoMNRB") = True Then
            FrmSendtoMNRB.SetFocus
        Else
            FrmSendtoMNRB.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuToPayor_Click()
On Error GoTo errOpen
    If CheckAccess(35) = True Then 'Set Delivery Date (Send to Payor)
        If FormLoadedByName("frmSendtoPayor") = True Then
            FrmSendtoPayor.SetFocus
        Else
            FrmSendtoPayor.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuBordxTracking_Click()
On Error GoTo errOpen
    If CheckAccess(42) = True Then 'Bordx & Worksheet
        If FormLoadedByName("frmBordxTracking") = True Then
            FrmBordxTracking.SetFocus
        Else
            FrmBordxTracking.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub ClaimTracking_Click()
If CheckTime("TRACKING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(44) = True Then 'Claim Tracking
        If FormLoadedByName("FrmClaimRpt") = True Then
            FrmClaimRpt.SetFocus
        Else
            FrmClaimRpt.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub ClaimTracking2_Click()
If CheckTime("TRACKING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(45) = True Then 'Claim Tracking 2
        If FormLoadedByName("FrmClaimRpt2") = True Then
            FrmClaimRpt2.SetFocus
        Else
            FrmClaimRpt2.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuWksTracking2_Click()
On Error GoTo errOpen
    If CheckAccess(68) = True Then 'Worksheet Tracking (Neptune)
        If FormLoadedByName("FrmWorksheetTracking2") = True Then
            FrmWorksheetTracking2.SetFocus
        Else
            FrmWorksheetTracking2.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuWSAdjGST_Click()
On Error GoTo errOpen
    If CheckAccess(65) = True Then 'Worksheet Adjustment
        If FormLoadedByName("frmWSAdjGST") = True Then
            frmWSAdjGST.SetFocus
        Else
            frmWSAdjGST.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuWSEnq_Click()
On Error GoTo errOpen
    If CheckAccess(13) = True Then 'Case Inquiry
        If FormLoadedByName("frmWorksheetEnq") = True Then
            frmWorksheetEnq.SetFocus
        Else
            frmWorksheetEnq.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuWSReturnCLM_Click()
On Error GoTo errOpen
    If CheckAccess(66) = True Then 'Worksheet Returned From Claims
        If FormLoadedByName("FrmARRecFrCLM") = True Then
            FrmARRecFrCLM.SetFocus
        Else
            FrmARRecFrCLM.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuWSToAC_Click()
On Error GoTo errOpen
    If CheckAccess(66) = True Then 'Worksheet Send to A/C
        If FormLoadedByName("frmWksSendToAC") = True Then
            FrmWksSendToAC.SetFocus
        Else
            FrmWksSendToAC.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub MnuZurichRejectPending_Click()
On Error GoTo errOpen
    If CheckAccess(18) = True Then 'Claim ASCII
        If FormLoadedByName("frmClaimASCIIRejZurich") = True Then
            frmClaimASCIIRejZurich.SetFocus
        Else
            frmClaimASCIIRejZurich.Show
        End If
    Else
        MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuZurichUpload_Click()
On Error GoTo errOpen
    If CheckAccess(8) = True Then 'Upload Bordx
        If FormLoadedByName("frmMBMUpload2015.show") = True Then
             frmMBMUpload2015.SetFocus
        Else
             frmMBMUpload2015.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
 Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub PaymentToHospital_Click()
On Error GoTo errOpen
    If CheckAccess(145) = True Then 'Enquiry Payment to Hospital
        If FormLoadedByName("frmEnquiryPaymentHospital") = True Then
            FrmEnquiryPaymentHospital.SetFocus
        Else
            FrmEnquiryPaymentHospital.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub PaymentToMember_Click()
On Error GoTo errOpen
    If CheckAccess(146) = True Then 'Enquiry Payment to Member
        If FormLoadedByName("frmEnquiryPayMember") = True Then
            FrmEnquiryPayMember.SetFocus
        Else
            FrmEnquiryPayMember.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub PayorExp_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(172) = True Then 'Posting - Receipt - Payor
        If FormLoadedByName("frmPostingRecPayor") = True Then
            frmPostingRecPayor.SetFocus
        Else
            frmPostingRecPayor.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub Post2MBM_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(174) = True Then 'Posting - Payment - MBM
        If FormLoadedByName("frmPostingPay2MBM") = True Then
            frmPostingPay2MBM.SetFocus
        Else
            frmPostingPay2MBM.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub PostPay2Hosp_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(173) = True Then 'Posting - Payment - HOS
        If FormLoadedByName("frmPostingPay2Hosp") = True Then
            frmPostingPay2Hosp.SetFocus
        Else
            frmPostingPay2Hosp.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub RECCLMFloatReq_Click()
On Error GoTo errOpen
    If CheckAccess(59) = True Then 'Receipt from Member
        If FormLoadedByName("frmCLMFloatTopupReq") = True Then
            frmCLMFloatTopupReq.SetFocus
        Else
            frmCLMFloatTopupReq.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub RECCLMFloatTopup_Click()
On Error GoTo errOpen
    If CheckAccess(59) = True Then 'Receipt from Member
        If FormLoadedByName("frmCLMFloatTopup") = True Then
            frmCLMFloatTopup.SetFocus
        Else
            frmCLMFloatTopup.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub RecCLMFloatTrans_Click()
On Error GoTo errOpen
    If CheckAccess(59) = True Then 'Receipt from Member
        If FormLoadedByName("frmCLMFloatTransfer") = True Then
            frmCLMFloatTransfer.SetFocus
        Else
            frmCLMFloatTransfer.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub RecExpOthers_Click()
If CheckTime("POSTING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(172) = True Then 'Posting - Receipt - Payor
        If FormLoadedByName("frmPostingRecOthers") = True Then
            frmPostingRecOthers.SetFocus
        Else
            frmPostingRecOthers.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub RetfrPay_Click()
On Error GoTo errOpen
    If CheckAccess(36) = True Then 'Set Delivery Date (Returned From Payor)
        If FormLoadedByName("frmReturnfrPayor") = True Then
            frmReturnFrPayor.SetFocus
        Else
            frmReturnFrPayor.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub StoDocument_Click()
On Error GoTo errOpen
    If CheckAccess(48) = True Then 'Storage
        If FormLoadedByName("frmStorageDoc") = True Then
            frmStorageDoc.SetFocus
        Else
            frmStorageDoc.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub StorageEnquiry_Click()
On Error GoTo errOpen
    If CheckAccess(49) = True Then 'Storage
        If FormLoadedByName("FrmStorageEnq") = True Then
            frmStorageEnq.SetFocus
        Else
            frmStorageEnq.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub StorageEntry_Click()
On Error GoTo errOpen
    If CheckAccess(48) = True Then 'Storage
        If FormLoadedByName("FrmStorageEntry") = True Then
            frmStorageEntry.SetFocus
        Else
            frmStorageEntry.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub StoRepEnq_Click()
On Error GoTo errOpen
    If CheckAccess(152) = True Then 'Storage
        If FormLoadedByName("frmStorageRepEnq") = True Then
            frmStorageRepEnq.SetFocus
        Else
            frmStorageRepEnq.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub StoRepudiateCase_Click()
On Error GoTo errOpen
    If CheckAccess(151) = True Then 'Storage Repudiate Cases
        If FormLoadedByName("FrmStorageRepEntry") = True Then
            FrmStorageRepEntry.SetFocus
        Else
            FrmStorageRepEntry.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub Suspension_Click()
On Error GoTo errOpen
    If CheckAccess(110) = True Then 'Case Inquiry
        If FormLoadedByName("frmSuspensionDetail") = True Then
            frmSuspensionDetail.SetFocus
        Else
            frmSuspensionDetail.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub UpdateBordx_Click()
On Error GoTo errOpen
    If CheckAccess(34) = True Then 'Update Bordx
        If FormLoadedByName("FrmUpdatews2bordx") = True Then
            frmUpdateWS2Bordx.SetFocus
        Else
            frmUpdateWS2Bordx.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub BorAccPeriod_Click()
On Error GoTo errOpen

If CheckAccess(39) = True Then 'Bordx A/C Period
        If FormLoadedByName("frmACPeriod") = True Then
            FrmACPeriod.SetFocus
        Else
            FrmACPeriod.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuGenerateBordx_Click()
On Error GoTo errOpen
    If CheckAccess(30) = True Then 'Insert Worksheet
        If FormLoadedByName("frmWksBordx") = True Then
            FrmWksBordx.SetFocus
        Else
            FrmWksBordx.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnuPrintBordx_Click()
On Error GoTo errOpen
    If CheckAccess(32) = True Then 'Print Bordx
        If FormLoadedByName("frmPrintBordx") = True Then
            FrmPrintBordx.SetFocus
        Else
            FrmPrintBordx.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub BordxTracking_Click()
On Error GoTo errOpen

If CheckAccess(43) = True Then 'Bordx Tracking
        If FormLoadedByName("frmBordxTracking2") = True Then
            FrmBordxTracking2.SetFocus
        Else
            FrmBordxTracking2.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuBordxCtrl_Click()
On Error GoTo errOpen
    If CheckAccess(40) = True Then 'Bordx Control
        If FormLoadedByName("frmBordxControl") = True Then
            FrmBordxControl.SetFocus
        Else
            FrmBordxControl.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnuBordxRevise_Click()
On Error GoTo errOpen
    If CheckAccess(31) = True Then 'Remove Worksheet
        If FormLoadedByName("frmBordxRevise") = True Then
            FrmBordxRevise.SetFocus
        Else
            FrmBordxRevise.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub
' ================= Claim Bordx ==================================

' ================= Case =========================================

Private Sub mnuLGTracking_Click()
On Error GoTo errOpen
If CheckAccess(17) = True Then 'LG Tracking
    If FormLoadedByName("frmLGTracking") = True Then
        frmLGTracking.SetFocus
    Else
        frmLGTracking.Show
    End If
Else
    MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
   
End Sub


Private Sub CaseAdjustment_Click()
On Error GoTo errOpen

If CheckAccess(12) = True Then 'Case Adjustment
        If FormLoadedByName("frmCaseAdjustment2016") = True Then
            frmCaseAdjustment2016.SetFocus
        Else
            frmCaseAdjustment2016.Show
            If CheckAccess(178) = True Then
                frmCaseAdjustment2016.cmdDelGL.Visible = True
                frmCaseAdjustment2016.cmdClearDoc.Visible = True
            Else
                frmCaseAdjustment2016.cmdDelGL.Visible = False
                frmCaseAdjustment2016.cmdClearDoc.Visible = False
            End If
            
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub CaseEntry_Click()
On Error GoTo errOpen
    
    If CheckAccess(11) = True Then 'Case Registration
        If FormLoadedByName("frmCaseAdmission") = True Then
            frmCaseAdmission.SetFocus
        Else
            frmCaseAdmission.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub CaseTracking_Click()
If CheckTime("TRACKING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(14) = True Then 'Case Tracking
        If FormLoadedByName("frmCaseTracking") = True Then
            FrmCaseTracking.SetFocus
        Else
            FrmCaseTracking.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub CASInMBM_Click()
On Error GoTo errOpen
    If CheckAccess(13) = True Then 'Case Inquiry
        If FormLoadedByName("frmCaseInquiry2016") = True Then
            frmCaseInquiry2016.SetFocus
        Else
            frmCaseInquiry2016.Show
            If CheckAccess(178) = True Then
                frmCaseInquiry2016.cmdPrintPymt.Enabled = True
            Else
                frmCaseInquiry2016.cmdPrintPymt.Enabled = False
            End If
            
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnuClaimAscii_Click()
On Error GoTo errOpen
If CheckAccess(18) = True Then 'Claim ASCII
    If FormLoadedByName("frmNewClaimAscii") = True Then
        frmNewClaimAscii.SetFocus
    Else
        frmNewClaimAscii.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

' ================= Case =========================================


' ================ Invoice ==============================
Private Sub mniInvoiceProcess_Click()
On Error GoTo errOpen
    If CheckAccess(22) = True Then 'Invoice Entry
    
        'If FormLoadedByName("frmInvoice") = True Then
        '    frmInvoice.SetFocus
        'Else
        '    frmInvoice.Show
        'End If
        
        If FormLoadedByName("frmInvoiceGST") = True Then
            frmInvoiceGST.SetFocus
        Else
            frmInvoiceGST.Show
        End If
        
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnuInvTrack_Click()
If CheckTime("TRACKING") = False Then Exit Sub
On Error GoTo errOpen
  If CheckAccess(27) = True Then 'Invoice Tracking
        If FormLoadedByName("frmInvoiceTracking2") = True Then
            FrmInvoiceTracking2.SetFocus
        Else
            FrmInvoiceTracking2.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
  Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuInvMBMAcc_Click()
On Error GoTo errOpen
  If CheckAccess(23) = True Then 'Invoice MBM Accounting
        If FormLoadedByName("frmInvoiceMBMAC") = True Then
            FrmInvoiceMBMAC.SetFocus
        Else
            FrmInvoiceMBMAC.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
  Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuInvHOSAcc_Click()
On Error GoTo errOpen
  If CheckAccess(24) = True Then 'Invoice HOS Accounting
        If FormLoadedByName("frmInvoiceHOSAC") = True Then
            FrmInvoiceHOSAC.SetFocus
        Else
            FrmInvoiceHOSAC.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
  Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub
' ================ Invoice =================================


' ===================== MCO Fee ===========================

Private Sub RefundListing_Click()
On Error GoTo errOpen
If CheckAccess(21) = True Then 'MCO Refund Listing
    If FormLoadedByName("frmMCORefundListing2") = True Then
        frmMCORefundListing2.SetFocus
    Else
        frmMCORefundListing2.Show
    End If
Else
    MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

' ===================== Worksheet ===========================

Private Sub mnuNewWorksheet_Click()
On Error GoTo errOpen
    If CheckAccess(64) = True Then 'Worksheet Registration
        'If FormLoadedByName("frmWorksheetReg2012v2") = True Then
        '    frmWorksheetReg2012v2.SetFocus
        'Else
        '    frmWorksheetReg2012v2.Show
        'End If
        If FormLoadedByName("frmWSRegGST2015V2") = True Then
            frmWSRegGST2015V2.SetFocus
        Else
            frmWSRegGST2015V2.Show
        End If
        
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnuWSAdjust_Click()
On Error GoTo errOpen
    If CheckAccess(65) = True Then 'Worksheet Adjustment
        'If FormLoadedByName("frmWorksheetAdj2013") = True Then
        '    frmWorksheetAdj2013.SetFocus
        'Else
        '    frmWorksheetAdj2013.Show
        'End If
        
        If FormLoadedByName("frmWSAdjGST") = True Then
            frmWSAdjGST.SetFocus
        Else
            frmWSAdjGST.Show
        End If
        
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuWksTracking_Click()
If CheckTime("TRACKING") = False Then Exit Sub
On Error GoTo errOpen
    If CheckAccess(67) = True Then 'Worksheet Tracking
        If FormLoadedByName("FrmWorksheetTracking") = True Then
            FrmWorksheetTracking.SetFocus
        Else
            FrmWorksheetTracking.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub
' =================== Worksheet ==============================

' ===================== Payment =========================

Private Sub mnuPaymentHos_Click()
On Error GoTo errOpen
    If CheckAccess(55) = True Then 'Payment to Hospital
        If FormLoadedByName("frmPaymentHospital") = True Then
            FrmPaymentHospital.SetFocus
        Else
            FrmPaymentHospital.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnuPaymentTracking_Click()
On Error GoTo errOpen
    If CheckAccess(58) = True Then 'Payment Tracking
        If FormLoadedByName("FrmPaymentTracking") = True Then
            FrmPaymentTracking.SetFocus
        Else
            FrmPaymentTracking.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub
' ======================= Payment ========================


' ====================== Maintainance ====================

Private Sub MaintenanceGeneral_Click()
   On Error GoTo errOpen
   Dim f As Form
   
   If (CheckAccess(69) = True Or CheckAccess(70) = True Or CheckAccess(71) = True Or _
       CheckAccess(72) = True Or CheckAccess(73) = True Or CheckAccess(74) = True Or _
       CheckAccess(75) = True Or CheckAccess(76) = True Or CheckAccess(77) = True Or _
       CheckAccess(78) = True Or CheckAccess(79) = True) Then
      
        If FormLoadedByName("frmMaintenanceGeneral") = True Then
            frmMaintenanceGeneral.SetFocus
        Else
            frmMaintenanceGeneral.Show
        End If

    Set f = frmMaintenanceGeneral
      If CheckAccess(69) = True Then
        f.SSTab1.TabEnabled(0) = True
      Else
        f.SSTab1.TabEnabled(0) = False
      End If
      
      If CheckAccess(70) = True Then
        f.SSTab1.TabEnabled(1) = True
      Else
        f.SSTab1.TabEnabled(1) = False
      End If
      
      If CheckAccess(71) = True Then
        f.SSTab1.TabEnabled(2) = True
      Else
        f.SSTab1.TabEnabled(2) = False
      End If
      
      If CheckAccess(72) = True Then
        f.SSTab1.TabEnabled(3) = True
      Else
        f.SSTab1.TabEnabled(3) = False
      End If
      
      If CheckAccess(73) = True Then
        f.SSTab1.TabEnabled(4) = True
      Else
        f.SSTab1.TabEnabled(4) = False
      End If
      
      If CheckAccess(74) = True Then
        f.SSTab1.TabEnabled(5) = True
      Else
        f.SSTab1.TabEnabled(5) = False
      End If
      
      If CheckAccess(75) = True Then
        f.SSTab1.TabEnabled(6) = True
      Else
        f.SSTab1.TabEnabled(6) = False
      End If
      If CheckAccess(76) = True Then
        f.SSTab1.TabEnabled(7) = True
      Else
        f.SSTab1.TabEnabled(7) = False
      End If
      If CheckAccess(77) = True Then
        f.SSTab1.TabEnabled(8) = True
      Else
        f.SSTab1.TabEnabled(8) = False
      End If
      If CheckAccess(78) = True Then
        f.SSTab1.TabEnabled(9) = True
      Else
        f.SSTab1.TabEnabled(9) = False
      End If
      If CheckAccess(79) = True Then
        f.SSTab1.TabEnabled(10) = True
      Else
        f.SSTab1.TabEnabled(10) = False
      End If
   Else
      MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
   End If
   Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
         
End Sub

Private Sub MaintenancePlan_Click()
On Error GoTo errOpen
    If (CheckAccess(80) = True Or CheckAccess(81) = True Or CheckAccess(82) = True Or _
    CheckAccess(83) = True Or CheckAccess(84) = True Or CheckAccess(85) = True) Then
      
        If FormLoadedByName("frmMaintenancePlan") = True Then
            frmMaintenancePlan.SetFocus
        Else
            frmMaintenancePlan.Show
        End If

      
      If CheckAccess(80) = True Then
        frmMaintenancePlan.SSTab1.TabEnabled(0) = True
      Else
        frmMaintenancePlan.SSTab1.TabEnabled(0) = False
      End If
      
      If CheckAccess(81) = True Then
        frmMaintenancePlan.SSTab1.TabEnabled(1) = True
      Else
        frmMaintenancePlan.SSTab1.TabEnabled(1) = False
      End If
      
      If CheckAccess(82) = True Then
        frmMaintenancePlan.SSTab1.TabEnabled(2) = True
      Else
        frmMaintenancePlan.SSTab1.TabEnabled(2) = False
      End If
      
      If CheckAccess(83) = True Then
        frmMaintenancePlan.SSTab1.TabEnabled(3) = True
      Else
        frmMaintenancePlan.SSTab1.TabEnabled(3) = False
      End If
      
      If CheckAccess(84) = True Then
        frmMaintenancePlan.SSTab1.TabEnabled(4) = True
      Else
        frmMaintenancePlan.SSTab1.TabEnabled(4) = False
      End If
      If CheckAccess(85) = True Then
        frmMaintenancePlan.SSTab1.TabEnabled(5) = True
      Else
        frmMaintenancePlan.SSTab1.TabEnabled(5) = False
      End If

   Else
      MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
   End If
   Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub


Private Sub mnuBenefits_Click()
    On Error GoTo errOpen
    If (CheckAccess(86) = True Or CheckAccess(87) = True _
    Or CheckAccess(88) = True Or CheckAccess(89) = True) Then
      
        If FormLoadedByName("frmMaintenanceBenefits") = True Then
            frmMaintenanceBenefits.SetFocus
        Else
            frmMaintenanceBenefits.Show
        End If

      
      If CheckAccess(86) = True Then
        frmMaintenanceBenefits.SSTab1.TabEnabled(0) = True
      Else
        frmMaintenanceBenefits.SSTab1.TabEnabled(0) = False
      End If
      
      If CheckAccess(87) = True Then
        frmMaintenanceBenefits.SSTab1.TabEnabled(1) = True
      Else
        frmMaintenanceBenefits.SSTab1.TabEnabled(1) = False
      End If
      
      If CheckAccess(88) = True Then
        frmMaintenanceBenefits.SSTab1.TabEnabled(2) = True
      Else
        frmMaintenanceBenefits.SSTab1.TabEnabled(2) = False
      End If
      
      If CheckAccess(89) = True Then
        frmMaintenanceBenefits.SSTab1.TabEnabled(3) = True
      Else
        frmMaintenanceBenefits.SSTab1.TabEnabled(3) = False
      End If
   Else
      MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
   End If
   Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub MaintenanceCase_Click()
On Error GoTo errOpen
If (CheckAccess(90) = True Or CheckAccess(91) = True Or CheckAccess(92) = True Or _
    CheckAccess(93) = True Or CheckAccess(94) = True Or CheckAccess(95) = True Or _
    CheckAccess(96) = True Or CheckAccess(97) = True) Then
    
    If FormLoadedByName("frmMaintenanceCase") = True Then
        frmMaintenanceCase.SetFocus
    Else
        frmMaintenanceCase.Show
    End If
    
    
    If CheckAccess(90) = True Then
        frmMaintenanceCase.SSTab1.TabEnabled(0) = True
    Else
        frmMaintenanceCase.SSTab1.TabEnabled(0) = False
    End If
    
    If CheckAccess(91) = True Then
        frmMaintenanceCase.SSTab1.TabEnabled(1) = True
    Else
        frmMaintenanceCase.SSTab1.TabEnabled(1) = False
    End If

    If CheckAccess(92) = True Then
        frmMaintenanceCase.SSTab1.TabEnabled(2) = True
    Else
        frmMaintenanceCase.SSTab1.TabEnabled(2) = False
    End If
    
    If CheckAccess(93) = True Then
        frmMaintenanceCase.SSTab1.TabEnabled(3) = True
    Else
        frmMaintenanceCase.SSTab1.TabEnabled(3) = False
    End If
    
    If CheckAccess(94) = True Then
        frmMaintenanceCase.SSTab1.TabEnabled(4) = True
    Else
        frmMaintenanceCase.SSTab1.TabEnabled(4) = False
    End If
    
    If CheckAccess(95) = True Then
        frmMaintenanceCase.SSTab1.TabEnabled(5) = True
    Else
        frmMaintenanceCase.SSTab1.TabEnabled(5) = False
    End If
    
    If CheckAccess(96) = True Then
        frmMaintenanceCase.SSTab1.TabEnabled(6) = True
    Else
        frmMaintenanceCase.SSTab1.TabEnabled(6) = False
    End If
    
    If CheckAccess(97) = True Then
        frmMaintenanceCase.SSTab1.TabEnabled(7) = True
    Else
        frmMaintenanceCase.SSTab1.TabEnabled(7) = False
    End If
Else
    MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
End If
Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub MnuMemberMaintenance_Click()
On Error GoTo errOpen
    If CheckAccess(109) = True Then 'Maintenance Date
        If FormLoadedByName("frmMaintenanceMembers") = True Then
            frmMaintenanceMembers.SetFocus
        Else
            frmMaintenanceMembers.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub MaintenanceOthers_Click()
On Error Resume Next
    
    Dim f As Form
    
    If (CheckAccess(98) = True Or CheckAccess(99) = True Or CheckAccess(100) = True Or _
       CheckAccess(101) = True Or CheckAccess(102) = True Or CheckAccess(103) = True Or _
       CheckAccess(104) = True Or CheckAccess(105) = True Or CheckAccess(106) = True Or _
       CheckAccess(107) = True Or CheckAccess(108) = True) Then
      
        If FormLoadedByName("frmMaintenanceOthers") = True Then
            frmMaintenanceOthers.SetFocus
        Else
            frmMaintenanceOthers.Show
        End If

      Set f = frmMaintenanceOthers
      If CheckAccess(98) = True Then
        f.SSTab1.TabEnabled(0) = True
      Else
        f.SSTab1.TabEnabled(0) = False
      End If
      
      If CheckAccess(99) = True Then
        f.SSTab1.TabEnabled(1) = True
      Else
        f.SSTab1.TabEnabled(1) = False
      End If
      
      If CheckAccess(100) = True Then
        f.SSTab1.TabEnabled(2) = True
      Else
        f.SSTab1.TabEnabled(2) = False
      End If
      
      If CheckAccess(101) = True Then
        f.SSTab1.TabEnabled(3) = True
      Else
        f.SSTab1.TabEnabled(3) = False
      End If
      
      If CheckAccess(102) = True Then
        f.SSTab1.TabEnabled(4) = True
      Else
        f.SSTab1.TabEnabled(4) = False
      End If
      
      If CheckAccess(103) = True Then
        f.SSTab1.TabEnabled(5) = True
      Else
        f.SSTab1.TabEnabled(5) = False
      End If
      
      If CheckAccess(104) = True Then
        f.SSTab1.TabEnabled(6) = True
      Else
        f.SSTab1.TabEnabled(6) = False
      End If
   
      If CheckAccess(105) = True Then
        f.SSTab1.TabEnabled(7) = True
      Else
        f.SSTab1.TabEnabled(7) = False
      End If
      If CheckAccess(106) = True Then
        f.SSTab1.TabEnabled(8) = True
      Else
        f.SSTab1.TabEnabled(8) = False
      End If
      
      If CheckAccess(107) = True Then
        f.SSTab1.TabEnabled(9) = True
      Else
        f.SSTab1.TabEnabled(9) = False
      End If
   
      If CheckAccess(108) = True Then
        f.SSTab1.TabEnabled(10) = True
      Else
        f.SSTab1.TabEnabled(10) = False
      End If
   Else
      MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
   End If
Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnuComInfo_Click()
On Error GoTo errOpen
    If CheckAccess(88) = True Then 'Company Info
        If FormLoadedByName("frmMaintenanceComInfo") = True Then
            FrmMaintenanceComInfo.SetFocus
        Else
            FrmMaintenanceComInfo.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub



Private Sub GroupMaintenance_Click()
On Error GoTo errOpen
    Dim f As New frmGroupMaintenance
    If CheckAccess(111) = True Then
        f.Show
    Else
        MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuPermission_Click()
On Error GoTo errOpen
If CheckAccess(112) = True Then 'User Permission
    If FormLoadedByName("frmUserPermissionMaintenance") = True Then
        frmUserPermissionMaintenance.SetFocus
    Else
        frmUserPermissionMaintenance.Show
    End If
Else
    MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
End If
Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub
' ====================== Maintainance ====================


' ==================== Receipt ==============================

Private Sub mnuReceiptMBM_Click()
On Error GoTo errOpen
    If CheckAccess(59) = True Then 'Receipt from Member
        If FormLoadedByName("frmReceiptMember") = True Then
            FrmReceiptMember.SetFocus
        Else
            FrmReceiptMember.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuReceipt_Click()
On Error GoTo errOpen
    If CheckAccess(62) = True Then 'Receipt from MNRB
        If FormLoadedByName("FrmReceiptMNRB") = True Then
            FrmReceiptMNRB.SetFocus
        Else
            FrmReceiptMNRB.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnuReceiptTracking_Click()
On Error GoTo errOpen
    
    If CheckAccess(63) = True Then 'Receipt Tracking
        If FormLoadedByName("frmReceiptTracking") = True Then
            FrmReceiptTracking.SetFocus
        Else
            FrmReceiptTracking.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub
' ==================== Receipt ============================

' ==================== Membership ============================

Private Sub mnuMBMAscii_Click()
On Error GoTo errOpen
    If CheckAccess(9) = True Then 'Generate MBM ASCII
        If FormLoadedByName("frmMBMAscii") = True Then
            FrmMBMAscii.SetFocus
        Else
            FrmMBMAscii.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
 Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub mnuUpload_Click()
On Error GoTo errOpen
    If CheckAccess(8) = True Then 'Upload Bordx
        If FormLoadedByName("frmMBMUpload2015.show") = True Then
        'If FormLoadedByName("FrmNewMBMEndt10.show") = True Then
             'FrmNewMBMEndt10.SetFocus
             frmMBMUpload2015.SetFocus
        Else
             'FrmNewMBMEndt10.Show
             frmMBMUpload2015.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
 Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub MembershipRegistration_Click()
On Error GoTo errOpen
    If CheckAccess(1) = True Then 'Membership Registration
        If FormLoadedByName("frmMemberEntry") = True Then
            frmMemberEntry.SetFocus
        Else
            frmMemberEntry.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub MembershipAdjustment_Click()
On Error GoTo errOpen
    
    If CheckAccess(2) = True Then 'Membership Adjustment
        If FormLoadedByName("frmNewMBMAdjustment") = True Then
            frmNewMBMAdjustment.SetFocus
        Else
            frmNewMBMAdjustment.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub MembershipDelete_Click()
On Error GoTo errOpen
   
    If CheckAccess(5) = True Then 'Membership Deletion
        If FormLoadedByName("frmDeleteMembers") = True Then
            frmDeleteMembers.SetFocus
        Else
            frmDeleteMembers.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub MembershipEnquiry_Click()
On Error GoTo errOpen
    If CheckAccess(3) = True Then 'Membership Inquiry
        If FormLoadedByName("frmNewMBMEnquiry") = True Then
            frmNewMBMEnquiry.SetFocus
        Else
            frmNewMBMEnquiry.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub MembershipReport_Click()
On Error GoTo errOpen
    If CheckAccess(4) = True Then 'Leaflet & Card
        If FormLoadedByName("frmReportMain") = True Then
            frmReportMain.SetFocus
        Else
            frmReportMain.Show
        End If
    Else
       MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub MembershipTracking_Click()
If CheckTime("TRACKING") = False Then Exit Sub
On Error GoTo errOpen
       
    If CheckAccess(6) = True Then 'Membership Tracking
        If FormLoadedByName("frmNewMBMTracking") = True Then
            FrmNewMBMTracking.SetFocus
        Else
            FrmNewMBMTracking.Show
        End If
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
    
End Sub



Private Function CheckTime(tmpAction As String) As Boolean
Dim TimeRec As New ADODB.Recordset
Dim tmpStartTime As String
Dim TimeSql As String
Dim tmpUserTime As String
Dim tmpTime As String

    'Call GetRemoteTime
    
    CheckTime = True
    If SpcAccess = "Y" Then Exit Function
    'tmptime = GetRemoteTime
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    TimeSql = "Select StartTime from TimeMain where Process ='" & Trim(tmpAction) & "'"
    TimeRec.Open TimeSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    Do While Not TimeRec.EOF
        tmpStartTime = TimeRec!StartTime
        TimeRec.MoveNext
    Loop
    TimeRec.Close
    Set TimeRec = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    
    If Trim(tmpStartTime) <> "" Then
        tmpUserTime = Format((Now), "h:m:ss")
        tmpStartTime = Format(tmpStartTime, "h:m:ss")
        tmpUserTime = CVar(tmpUserTime)
        tmpStartTime = CVar(tmpStartTime)
        If DateDiff("s", tmpUserTime, tmpStartTime) >= 0 Then
            MsgBox tmpAction & " is available at  " & Format(tmpStartTime, "hh:mm:ss AMPM") & " ! "
            CheckTime = False
        End If
    End If
End Function

Private Sub ZurichCardPrinting_Click()
On Error GoTo errOpen
    If CheckAccess(4) = True Then 'Leaflet & Card
        If FormLoadedByName("frmZurichCardPrinting") = True Then
            frmZurichCardPrinting.SetFocus
        Else
            frmZurichCardPrinting.Show
        End If
    Else
       MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If
End Sub

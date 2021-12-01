Attribute VB_Name = "basLocalizationConst"
Option Explicit

'
'*********************************************************************************
' This file contains the String and Message constant definitions for this project
' created by the SDK Wizard.
'
'*********************************************************************************
'
' Strings
'
Public Const kNone = 103                                    ' (none)
Public Const kAPInvoiceNo = 6408                            ' Invoice No

'
' Messages
'
Public Const kmsgCannotBeBlank = 130082                     ' {0} cannot be blank.
Public Const kmsgVoucherPurchAmtNeg = 140044                ' A {0} may not have a Purchases Amount less than zero!The current Purchases Amount Total is:  {1}Please adjust one or more detail lines so that the Purchases Amount is greater than or equal to zero.
Public Const kmsgVoucherInvTotalNeg = 140043                ' A {0} may not have an Invoice Total less than zero!The current Invoice Total is:  {1}Please adjust the detail line(s), Sales Tax amount or Freight amount so that the Invoice Total is greater than or equal to zero.
Public Const kmsgVoucherOutOfBalance = 140005               ' This {0} is out of balance!The current Undistributed Balance is:  {1}Do you want the system to automatically adjust the Invoice Amount for you?
Public Const kmsgVoucherTermsDiscNeg = 140065               ' A {0} may not have a Terms Discount Amount less than zero!Please enter a valid Terms Discount Amount.
Public Const kmsgSetCurrControlsError = 110098              ' The controls on this form which accept currency values cannot be set based on the parameters for the following Currency:Currency ID: {0}
Public Const kmsgInvTotalCannotBeZero = 140320              ' The Invoice Total cannot be zero for a {0} that is applied to a Check.
Public Const kmsgVoucherLogError = 140012                   ' An error occurred while attempting to update the {0} log.Please call Product Support for technical assistance.
Public Const kmsgUnexpectedConfirmUnloadRV = 100009         ' Unexpected Confirm Unload Return Value: {0}
Public Const kmsgDDNNotRegistered = 220056                  ' The drill down "{0}" is not registered on your machine. Please contact your system administrator.
Public Const kmsgSetupSelectionGridFail = 153368            ' Setup of the Selection Grid Failed.
Public Const kmsgProc = 153223                              ' Unexpected Error running Stored Procedure {0}.
Public Const kmsgEntryRequired = 100078                     ' You must specify a value for {0}
Public Const kmsgUnexpectedCtrlArryIndex = 100007           ' Unexpected Control Array Index Value
Public Const kmsgCantLoadDDData = 153369                    ' Unable to load data from Data Dictionary.
Public Const kmsgFatalReportInit = 153302                   ' Unexpected Error Initializing Report Engine.
Public Const kmsgFataSortInit = 153370                      ' Unable to initialize Sort Grid.
Public Const kmsgDocNoSave = 153464                         ' Warning: Exchange Rate is zero
Public Const kmsgARBadField = 150003                        ' You have entered an invalid value in {0}.  It will be changed back.
Public Const ksPer = 151381                                 ' per
Public Const kmsgNoCompanyID = 153498                       ' No CompanyID Set!  M/C Options could not be retrieved.
Public Const kmsgARNoRecalcOnExchRate = 153406              ' You have changed your currency or exchange rate. No amounts will automatically recalculate.


Attribute VB_Name = "basLocalizationConst"
Option Explicit

'
'*********************************************************************************
' This file contains the String and Message constant definitions for this project
' based upon the LocalString and LocalMessage tables.
'
' Created On: 1/7/2004 1:53:04 PM
'*********************************************************************************
'
' Strings
'
Public Const ksEnterPO = 220153                             ' Enter Purchase &Orders
Public Const ksEnterReq = 220154                            ' Enter &Requisitions
Public Const kSotaTitle = 1100                              ' MAS 500
Public Const kVendClassListing = 140159                     ' Vendor Class Listing
'
' Messages
'
Public Const kmsgSotaErr = 1002                             '             Module: {0}  Company:  {1}  AppName:  {2}Error < {3} >  occurred at < {4} >  in Procedure < {5} >{6}                                    This application will now close.
Public Const kmsgUnexpectedKeyChangeCode = 100008           ' Unexpected KeyChangeCode: {0}
Public Const kmsgUnexpectedConfirmUnloadRV = 100009         ' Unexpected Confirm Unload Return Value: {0}
Public Const kmsgRequiredField = 100012                     ' {0} required.
Public Const kmsgInvalidDate = 100024                       ' Invalid date entered
Public Const kmsgMustBeGreatZero = 100026                   ' Must enter a value greater than zero
Public Const kmsgUnexpectedSPReturnValue = 100071           ' Unexpected Stored Procedure Return Value: {0}
Public Const kmsgDMNoSaveConcurrency = 100231               ' This record has been modified by another user.  Your changes will not be saved.
Public Const kmsgPOIncomplete = 220028                      ' This PO exists in the log with a status of Incomplete another user may have created it.  Do you want to use it?
Public Const kmsgPOPurged = 220029                          ' This Purchase Order has been purged, you may not use this number.
Public Const kmsgPOLogBad = 220030                          ' The Purchase order log says there is an existing PO but one was not found.  Do you want to use this number?
Public Const kmsgAllPONumsUsed = 220036                     ' All PO numbers have been used
Public Const kmsgReqSaveChanges = 220067                    ' Save changes to {0}?
Public Const kmsgCannotDelReqLine = 220068                  ' The selected Requisition line cannot be deleted because it is associated with a Purchase Order.
Public Const kmsgReqGenSuccessful = 220069                  ' All of the selected lines were sucessfully generated into Purchase Orders.
Public Const kmsgReqGenPartialSuccessful = 220070           ' Some lines were not generated into Purchase Orders due to problems encountered
Public Const kmsgReqGenErrorEncountered = 220071            ' Errors encountered!!  No Purchase Orders were generated.
Public Const kmsgSPBadReturn = 210017                       ' Stored Procedure '{0}' returned an unexpected value: {1}
Public Const kmsgExpError = 153450                          ' Error: "{0}" in {1}
Public Const kmsgARBadField = 150003                        ' You have entered an invalid value in {0}.  It will be changed back.
Public Const kmsgNoBlankItemWithWhse = 220096               ' Item cannot be blank when associated with a Warehouse.
Public Const kmsgItemInvalidForWhse = 220098                ' The item specified is not valid for the selected warehouse.  It will be changed back.
Public Const kmsgItemInvalidForWhseDflt = 220099            ' The item specified is not valid for the default warehouse.  It will be changed back.
Public Const kmsgItemNotvalidForWhse = 220100               ' The item is not valid for the warehouse chosen.  Clear the item?
Public Const kmsgCannotBeGreaterThan = 120017               ' {0} cannot be greater than {1}
Public Const kmsgCannotBeNegative = 120019                  ' {0} cannot be a negative value.
Public Const kmsgCannotBeBlank = 130082                     ' {0} cannot be blank.
Public Const kmsgPOMissingWhseOrDept = 220173               ' Must Enter a Whse OR a Department for an Inventory type item.

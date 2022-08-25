' GPLineItem Class Module

Private c_dName As String

Private c_dAdd1 As String

Private c_dCity As String

Private c_dState As String

Private c_dZip As String

Private c_debitMemoNum As String

Private c_euLoc As String

Private c_euLocName As String

Private c_euAdd1 As String

Private c_euAdd2 As String

Private c_euCity As String

Private c_euState As String

Private c_euZip As String

Private c_invoiceNum As Long

Private c_invoiceDate As Date

Private c_invoiceLineItem As Long

Private c_gpSku As String

Private c_dItemNum As String

Private c_numCases As Long

Private c_toPrice As Currency

Private c_rebate As Currency

Private c_extendedRebate As Currency

 

Public Sub InitiateClassFields(dName As String, dAdd1 As String, dCity As String, dState As String, _

                                dZip As String, debitMemoNum As String, euLoc As String, euLocName As String, _

                                euAdd1 As String, euAdd2 As String, euCity As String, euState As String, _

                                euZip As String, invoiceNum As Long, invoiceDate As Date, invoiceLineItem As Long, _

                                gpSku As String, dItemNum As String, numCases As Long, toPrice As Currency, _

                                rebate As Currency, extendedRebate As Currency)

 

    c_dName = Left(dName, 30)

    c_dAdd1 = Left(dAdd1, 30)

    c_dCity = Left(dCity, 30)

    c_dState = dState

    c_dZip = dZip

    c_debitMemoNum = Left(debitMemoNum, 15)

   c_euLoc = Left(euLoc, 30)

    c_euLocName = Left(euLocName, 30)

    c_euAdd1 = Left(euAdd1, 30)

    c_euAdd2 = Left(euAdd2, 30)

    c_euCity = Left(euCity, 30)

    c_euState = euState

    c_euZip = Left(euZip, 10)

    c_invoiceNum = invoiceNum

    c_invoiceDate = invoiceDate

    c_invoiceLineItem = invoiceLineItem

    c_gpSku = Left(gpSku, 19)

    c_dItemNum = Left(dItemNum, 22)

    c_numCases = numCases

    c_toPrice = toPrice

    c_rebate = rebate

    c_extendedRebate = extendedRebate

   

    

End Sub

 

'Getters

Property Get getDName() As String

    getDName = c_dName

End Property

 

Property Get getDAdd1() As String

    getDAdd1 = c_dAdd1

End Property

 

Property Get getDCity() As String

    getDCity = c_dCity

End Property

 

 

Property Get getDState() As String

    getDState = c_dState

End Property

 

Property Get getDZip() As String

    getDZip = c_dZip

End Property

 

Property Get getDebitMemoNum() As String

    getDebitMemoNum = c_debitMemoNum

End Property

 

Property Get getEULoc() As String

    getEULoc = c_euLoc

End Property

 

Property Get getEULocName() As String

    getEULocName = c_euLocName

End Property

 

Property Get getEUAdd1() As String

    getEUAdd1 = c_euAdd1

End Property

 

Property Get getEUAdd2() As String

    getEUAdd2 = c_euAdd2

End Property

 

Property Get getEUCity() As String

    getEUCity = c_euCity

End Property

 

Property Get getEUState() As String

    getEUState = c_euState

End Property

 

Property Get getEUZip() As String

    getEUZip = c_euZip

End Property

 

Property Get getInvoiceNum() As Long

    getInvoiceNum = c_invoiceNum

End Property

 

Property Get getInvoiceDate() As Date

    getInvoiceDate = c_invoiceDate

End Property

 

Property Get getInvoiceLineItem() As Long

    getInvoiceLineItem = c_invoiceLineItem

End Property

 

Property Get getGPSku() As String

    getGPSku = c_gpSku

End Property

 

Property Get getDItemNum() As String

    getDItemNum = c_dItemNum

End Property

 

Property Get getNumCases() As Long

    getNumCases = c_numCases

End Property

 

Property Get getToPrice() As Currency

   getToPrice = c_toPrice

End Property

 

Property Get getRebate() As Currency

    getRebate = c_rebate

End Property

 

Property Get getExtendedRebate() As Currency

    getExtendedRebate = c_extendedRebate

End Property
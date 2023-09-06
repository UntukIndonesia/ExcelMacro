Dim alamat_database As String
Dim rowInputProjectName As Integer
Dim colInputProjectName As Integer
Dim rowInputPIC As Integer
Dim colInputPIC As Integer
Dim rowInputCategory As Integer
Dim colInputCategory As Integer
Dim rowInputContractType As Integer
Dim colInputContractType As Integer
Dim rowInputVendorName As Integer
Dim colInputVendorName As Integer
Dim rowInputNoKontrak As Integer
Dim colInputNoKontrak As Integer

Dim rowInputSetupNoKontrak As Integer
Dim colInputSetupNoKontrak As Integer

Dim rowDeleteNoKontrak As Integer
Dim colDeleteNoKontrak As Integer
Dim rowDeleteKeterangan As Integer
Dim colDeleteKeterangan As Integer

Dim rowDBStart As Integer
Dim colDBTanggal As Integer
Dim colDBNoKontrak As Integer
Dim colDBProjectName As Integer
Dim colDBPIC As Integer
Dim colDBContractType As Integer
Dim colDBCategoryType As Integer
Dim colDBVendorName As Integer
Dim colDBCancelRemark As Integer
Dim colDBNoKontrak2 As Integer

Sub INITIATION()
alamat_database = "C:\Users\Markus\Documents\Database Kontrak.xlsx"
rowInputProjectName = 3
colInputProjectName = 2
rowInputPIC = 4
colInputPIC = 2
rowInputCategory = 5
colInputCategory = 2
rowInputContractType = 6
colInputContractType = 2
rowInputVendorName = 7
colInputVendorName = 2
rowInputNoKontrak = 8
colInputNoKontrak = 2

rowInputSetupNoKontrak = 23
colInputSetupNoKontrak = 2

rowDeleteNoKontrak = 17
colDeleteNoKontrak = 2
rowDeleteKeterangan = 18
colDeleteKeterangan = 2

rowDBStart = 2
colDBTanggal = 1
colDBNoKontrak = 2
colDBProjectName = 3
colDBPIC = 4
colDBContractType = 5
colDBVendorName = 6
colDBCancelRemark = 8
colDBNoKontrak2 = 9
colDBCategoryType = 7
End Sub

Sub InputData()

Application.ScreenUpdating = False

INITIATION

Dim rowDB As Integer
Dim ProjectName As String
Dim PIC As String
Dim Category As String
Dim ContractType As String
Dim VendorName As String
Dim NoKontrak As String
Dim regEx As New VBScript_RegExp_55.RegExp
Dim matches, s
regEx.Pattern = "\d+\/(FIN)\/(MAIN|ADD)\/\d{2,2}\/\d{4,4}"
regEx.IgnoreCase = True 'True to ignore case
regEx.Global = True 'True matches all occurances, False matches the first occurance

Dim Temp_Kontrak As Integer
Dim Kontrak As String
Dim No_Kontrak As String
Dim Full_No_Kontrak As String

Dim Categories As String
Dim Contract_Type As String
Dim Months As String
Dim Years As Integer

'OPEN SHARED KONTRAK NO DATABASE
'AMBIL NAMA FILE MACRONYA
AppsMacro = ActiveWorkbook.Name

Workbooks.Open Filename:=alamat_database
NoKontrakMacro = ActiveWorkbook.Name
Range("A1").Select

If ActiveCell.Value <> "" Then

Workbooks(NoKontrakMacro).Activate

'Generate Number
Sheets("Database_Keseluruhan").Select
Range("B2").Select
Do While ActiveCell.Value <> ""
    ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Offset(-1, 0).Select
Temp_Kontrak = InStr(1, ActiveCell.Value, "/")
Kontrak = ActiveCell.Value
If Temp_Kontrak = 0 Then
    No_Kontrak = 1
Else
    No_Kontrak = Left(Kontrak, Temp_Kontrak - 1) + 1
End If

Workbooks(AppsMacro).Activate

Sheets("Input").Select
Categories = "FIN"
If Range("B6").Value = "Addendum" Then
    Contract_Type = "ADD"
Else
    Contract_Type = "MAIN"
End If
Months = Month(Now())
If Months < 10 Then
    Months = "0" & Months
Else
End If
Years = Year(Now())

Full_No_Kontrak = No_Kontrak & "/" & Categories & "/" & Contract_Type & "/" & Months & "/" & Years
Range("B8").Value = Full_No_Kontrak

Workbooks(NoKontrakMacro).Activate

Workbooks(NoKontrakMacro).Save

Workbooks(AppsMacro).Activate

Sheets("Input").Select

'VALIDASI INPUT
Dim ErrMsg As String
Dim ErrStatus As Integer
ErrStatus = 0
ErrMsg = "Error dalam memasukan data: "
With Sheets("Input")
  If .Cells(rowInputProjectName, colInputProjectName) = "" Then
    ErrMsg = ErrMsg + vbNewLine + " Nama Project Kosong"
    ErrStatus = -1
  End If
  If .Cells(rowInputPIC, colInputPIC) = "" Then
    ErrMsg = ErrMsg + vbNewLine + " PIC Kosong"
    ErrStatus = -1
  End If
  If .Cells(rowInputCategory, colInputCategory) = "" Then
    ErrMsg = ErrMsg + vbNewLine + " Kategori Kosong"
    ErrStatus = -1
  End If
  If .Cells(rowInputContractType, colInputContractType) = "" Then
    ErrMsg = ErrMsg + vbNewLine + " Tipe Kontrak Kosong"
    ErrStatus = -1
  End If
  If .Cells(rowInputVendorName, colInputVendorName) = "" Then
    ErrMsg = ErrMsg + vbNewLine + " Nama Vendor Kosong"
    ErrStatus = -1
  End If
  If .Cells(rowInputNoKontrak, colInputNoKontrak) = "" Then
    ErrMsg = ErrMsg + vbNewLine + " Nomor Kontrak Kosong"
    ErrStatus = -1
  End If
  Set allMatches = regEx.Execute(.Cells(rowInputNoKontrak, colInputNoKontrak))
  If allMatches.Count = 0 Then
    ErrMsg = ErrMsg + vbNewLine + " Nomor Kontrak Tidak sesuai format"
    ErrStatus = -1
  End If
End With

If ErrStatus <> -1 Then

    With Sheets("Input")
        ProjectName = .Cells(rowInputProjectName, colInputProjectName)
        PIC = .Cells(rowInputPIC, colInputPIC)
        Category = .Cells(rowInputCategory, colInputCategory)
        ContractType = .Cells(rowInputContractType, colInputContractType)
        VendorName = .Cells(rowInputVendorName, colInputVendorName)
        NoKontrak = .Cells(rowInputNoKontrak, colInputNoKontrak)
    End With
    
    
    With Sheets("Database_Lokal")
        'Membuka Proteksi
        .Unprotect Password:="Eden"
        rowDB = rowDBStart
        Do While .Cells(rowDB - 1, colDBNoKontrak) <> ""
            If .Cells(rowDB, colDBNoKontrak) = "" Then
                .Cells(rowDB, colDBProjectName) = ProjectName
                .Cells(rowDB, colDBNoKontrak) = NoKontrak
                .Cells(rowDB, colDBPIC) = PIC
                .Cells(rowDB, colDBContractType) = ContractType
                .Cells(rowDB, colDBVendorName) = VendorName
                .Cells(rowDB, colDBTanggal) = Now()
                .Cells(rowDB, colDBCategoryType) = Category
            Exit Do
            End If
            rowDB = rowDB + 1
            
        Loop
        .Protect Password:="Eden"
    End With
    
    Workbooks(NoKontrakMacro).Activate
    With Sheets("Database_Keseluruhan")
        .Unprotect Password:="Eden"
        rowDB = rowDBStart
        Do While .Cells(rowDB - 1, colDBNoKontrak) <> ""
            If .Cells(rowDB, colDBNoKontrak) = "" Then
                .Cells(rowDB, colDBProjectName) = ProjectName
                .Cells(rowDB, colDBNoKontrak) = NoKontrak
                .Cells(rowDB, colDBPIC) = PIC
                .Cells(rowDB, colDBContractType) = ContractType
                .Cells(rowDB, colDBVendorName) = VendorName
                .Cells(rowDB, colDBTanggal) = Now()
                .Cells(rowDB, colDBCategoryType) = Category
            Exit Do
            End If
            rowDB = rowDB + 1
        Loop
        .Protect Password:="Eden"
    End With
    Workbooks(NoKontrakMacro).Save
    Workbooks(NoKontrakMacro).Close
    Workbooks(AppsMacro).Activate
    
    Range("B3").Value = ""
    Range("B4").Value = ""
    Range("B5").Value = ""
    Range("B6").Value = ""
    Range("B7").Value = ""
    Range("B15").Value = ""
    Range("B16").Value = ""
    Workbooks(AppsMacro).Save
    
    
    
    MsgBox ("Pengisian Data Selesai")
    Sheets("Database_Lokal").Select
    
Else
    MsgBox (ErrMsg)
End If

Else
    
    Workbooks(AppsMacro).Activate
    
    'Ambil data
    With Sheets("Input")
        ProjectName = .Cells(rowInputProjectName, colInputProjectName)
        PIC = .Cells(rowInputPIC, colInputPIC)
        Category = .Cells(rowInputCategory, colInputCategory)
        ContractType = .Cells(rowInputContractType, colInputContractType)
        VendorName = .Cells(rowInputVendorName, colInputVendorName)
        Months = Month(Now())
        If Months < 10 Then
            Months = "0" & Months
        Else
        End If
        Years = Year(Now())
        'Generate Nomor kontrak
        NoKontrak = 1 & "/" & Category & "/" & ContractType & "/" & Months & "/" & Years
    End With
    
    'Buka sheet yang ingin dituju
    Workbooks.Open alamat_database
    
    'Buat Sheet
    Workbooks(NoKontrakMacro).Sheets.Add After:=Sheets(1)
    ActiveSheet.Name = "Database_Keseluruhan"
    
    'Buat Header
    With Sheets("Database_Keseluruhan")
        .Cells(1, 1) = "Tanggal"
        .Cells(1, 2) = "No Kontrak"
        .Cells(1, 3) = "Nama Project"
        .Cells(1, 4) = "PIC"
        .Cells(1, 5) = "Tipe Kontrak"
        .Cells(1, 6) = "Vendor Name"
        .Cells(1, 7) = "Remark"
        .Cells(1, 9) = "No Kontrak yang sudah digenerate"
    End With
    
    
    
    'Tempel data
    With Sheets("Database_Keseluruhan")
        rowDB = rowDBStart
        Do While .Cells(rowDB - 1, colDBNoKontrak) <> ""
            If .Cells(rowDB, colDBNoKontrak) = "" Then
                .Cells(rowDB, colDBProjectName) = ProjectName
                .Cells(rowDB, colDBNoKontrak) = NoKontrak
                .Cells(rowDB, colDBPIC) = PIC
                .Cells(rowDB, colDBContractType) = ContractType
                .Cells(rowDB, colDBVendorName) = VendorName
                .Cells(rowDB, colDBTanggal) = Now()
            Exit Do
            End If
            rowDB = rowDB + 1
        Loop
    End With
    
    MsgBox ("Path Database Nomor Kontrak Kosong")
    Workbooks(NoKontrakMacro).Save
    Workbooks(NoKontrakMacro).Close
End If

End Sub

Sub DeleteData()
INITIATION
Dim matching As String

    'AMBIL NAMA FILE MACRONYA
    AppsMacro = ActiveWorkbook.Name
        
        'ISI CODE LU DISINI BRO UNTUK YANG DELETE
        Workbooks(AppsMacro).Activate
        No_Kontrak = Range("B15").Value
        Keterangans = Range("B16").Value
        
        With Sheets("Database_Lokal")
        .Select
        .Unprotect Password:="Eden"
        End With
        
        Range("B2").Select
        Do While ActiveCell.Value <> ""
            If ActiveCell.Value = No_Kontrak Then
                ActiveCell.Offset(0, 6).Value = Keterangans
                matching = "Yes"
            Else
            End If
            ActiveCell.Offset(1, 0).Select
        Loop
        
        Workbooks(AppsMacro).Save
        
        With Sheets("Database_Lokal")
        .Protect Password:="Eden"
        End With
        
        Workbooks.Open Filename:=alamat_database
        NoKontrakMacro = ActiveWorkbook.Name
        
        With Sheets("Database_Keseluruhan")
        .Unprotect Password:="Eden"
        .Select
        End With
        
        Range("B2").Select
        Do While ActiveCell.Value <> ""
            If ActiveCell.Value = No_Kontrak Then
                ActiveCell.Offset(0, 6).Value = Keterangans
                matching = "Yes"
            Else
            End If
            ActiveCell.Offset(1, 0).Select
        Loop
        With Sheets("Database_Keseluruhan")
        .Protect Password:="Eden"
        End With
        Workbooks(NoKontrakMacro).Save
        Workbooks(NoKontrakMacro).Close
        
        Workbooks(AppsMacro).Activate
        Sheets("Input").Select
        If matching = "Yes" Then
            MsgBox ("Delete Data Selesai")
            Range("B16").Value = ""
        Else
            MsgBox ("Nomor kontrak tidak ditemukan")
            Range("B16").Value = ""
        End If
    
End Sub

Sub GenerateNumber()

Application.ScreenUpdating = False

Dim Temp_Kontrak As Integer
Dim Kontrak As String
Dim No_Kontrak As String
Dim Full_No_Kontrak As String

Dim Category As String
Dim Contract_Type As String
Dim Months As String
Dim Years As Integer

Sheets("Database").Select
Range("I2").Select
Do While ActiveCell.Value <> ""
    ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Offset(-1, 0).Select
Temp_Kontrak = InStr(1, ActiveCell.Value, "/")
Kontrak = ActiveCell.Value
No_Kontrak = Left(Kontrak, Temp_Kontrak - 1) + 1

Sheets("Input").Select
Category = Range("B5").Value
Contract_Type = Range("B6").Value
Months = Month(Now())
If Months < 10 Then
    Months = "0" & Months
Else
End If
Years = Year(Now())

Full_No_Kontrak = No_Kontrak & "/" & Category & "/" & Contract_Type & "/" & Months & "/" & Years
Range("B8").Value = Full_No_Kontrak

Sheets("Database").Select
Range("I2").Select

Do While ActiveCell.Value <> ""
    ActiveCell.Offset(1, 0).Select
Loop

ActiveCell.Value = Full_No_Kontrak

ThisWorkbook.Save

Sheets("Input").Select

End Sub
Sub SetupFolder()
INITIATION
sfile = Application.GetOpenFilename()
If sfile = False Then
    Exit Sub
End If

With Sheets("Input")
    .Cells(rowInputSetupNoKontrak, colInputSetupNoKontrak) = sfile
End With

End Sub

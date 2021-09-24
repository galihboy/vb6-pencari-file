VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplikasi Pencarian eBook - Galih Hermawan (Juli 2021)"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Shortcut"
      Height          =   735
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Width           =   2415
      Begin VB.ComboBox cboDir 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cari Data"
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   7935
      Begin VB.ListBox lstDaftar 
         Height          =   2010
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   7575
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         Height          =   255
         Left            =   5640
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtCari 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label lblHasilCari 
         Caption         =   "Jumlah buku hasil pencarian."
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   3495
      End
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   5655
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Pilih direktori :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Galih Hermawan
'galih.hermawan@gmail.com
'https://galih.eu

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

Private Const SB_BOTH = 3
Private Const LB_SETHORIZONTALEXTENT = &H194

' direktori pilihan untuk shortcut
Const dirPython = "D:\Buku\Python\_new\"
Const dirML1 = "D:\Buku\AI & Machine Learning\"
Const dirML2 = "D:\Buku\AI & Machine Learning\Machine Learning\"
Const dirDL = "D:\Buku\AI & Machine Learning\Deep Learning\"
Const dirNN = "D:\Buku\AI & Machine Learning\Neural Network\"
Const dirAlgoMath = "D:\Buku\Algoritma n Math\"

Dim arrDir(6, 2) As String

Private Sub cboDir_Click()
    ' jika combobox shortcut direktori diklik
    Dir1.path = arrDir(cboDir.ListIndex + 1, 2)

End Sub

Private Sub File1_DblClick()
    Dim almfile As String
    almfile = Dir1.path & "\" & File1.List(File1.ListIndex)
    'Buka file terpilih di program default
    ShellExecute hWnd, "open", almfile, vbNullString, vbNullString, 1
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    ' BEGIN - Menampilkan scrollbar horisontal listbox
    Dim lLength As Long
    lLength = 2 * (lstDaftar.Width / Screen.TwipsPerPixelX)
    Call SendMessage(lstDaftar.hWnd, LB_SETHORIZONTALEXTENT, lLength, 0&)
    ' END
    
    cboDir.Clear
    ' data nama dan alamat direktori shorcut dimasukkan ke array
    arrDir(1, 1) = "Python"
    arrDir(1, 2) = dirPython
    arrDir(2, 1) = "Machine Learning"
    arrDir(2, 2) = dirML1
    arrDir(3, 1) = "Machine Learning 2"
    arrDir(3, 2) = dirML2
    arrDir(4, 1) = "Deep Learning"
    arrDir(4, 2) = dirDL
    arrDir(5, 1) = "Neural Network"
    arrDir(5, 2) = dirNN
    arrDir(6, 1) = "Algo & Math"
    arrDir(6, 2) = dirAlgoMath
    ' data array dimasukkan ke combobox
    For i = 1 To UBound(arrDir)
        cboDir.AddItem arrDir(i, 1)
    Next
End Sub

Private Sub cmdCari_Click()
    
    Dim strLuaran() As String
    Dim i As Integer
    
    lstDaftar.Clear
    lstDaftar.Enabled = True
    ' panggil rutin CariData dan simpan hasilnya di strLuaran
    strLuaran = CariData(txtCari.Text, File1)
    
    ' Jika data array tidak kosong atau data ditemukan
    If validasiArray(strLuaran) Then
        For i = 1 To UBound(strLuaran)
            ' setiap data yang ditemukan, diletakkan di listbox
            lstDaftar.AddItem (strLuaran(i))
        Next
        lblHasilCari = "Ditemukan " & UBound(strLuaran) & " buku."
    Else
        lstDaftar.AddItem " ** Data tidak ditemukan"
        lstDaftar.Enabled = False
        lblHasilCari = "Buku tidak ditemukan."
    End If
    
End Sub

Private Sub Dir1_Change()
    ' jika direktori diubah
    File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
    ' jika drive (root directory) diubah
    Dir1.path = Left$(Drive1.Drive, 1) & ":\"
End Sub

' fungsi untuk mencari data
Function CariData(strCari As String, strInput As FileListBox) As String()
    Dim i As Integer
    Dim j As Integer
    Dim jmlData As Integer
    Dim luaran() As String
    Dim strSumber As String
    
    jmlData = strInput.ListCount
    j = 1
    For i = 0 To jmlData - 1
        strSumber = strInput.List(i)
        If CariTeks(LCase(strSumber), LCase(strCari)) Then
            ReDim Preserve luaran(j)
            luaran(j) = strSumber
            j = j + 1
        End If
    Next
    CariData = luaran
End Function

' fungsi untuk mencocokkan teks yang dicari dan teks (alamat file) tempat pencarian
Function CariTeks(strSumber As String, strCari As String) As Boolean
    If InStr(strSumber, strCari) <> 0 Then
        CariTeks = True
    Else
        CariTeks = False
    End If
End Function

' pemeriksaan array kosong atau tidak
Function validasiArray(arr() As String) As Boolean
    On Error GoTo ReturnFalse
    validasiArray = UBound(arr) >= LBound(arr)
ReturnFalse:
End Function

Private Sub lstDaftar_DblClick()
    Dim almfile As String
    almfile = Dir1.path & "\" & lstDaftar.List(lstDaftar.ListIndex)
    'Buka file terpilih di program default
    ShellExecute hWnd, "open", almfile, vbNullString, vbNullString, 1
End Sub

Private Sub txtCari_Change()
    cmdCari_Click
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00004000&
   Caption         =   "DiscoList Tool by Dragokas"
   ClientHeight    =   5055
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5055
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Text            =   "dragokas.com/music/index.php?id="
      Top             =   3720
      Width           =   4215
   End
   Begin VB.CheckBox chkSort 
      BackColor       =   &H00004000&
      Caption         =   "Сортировать"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton cmdCopyFiles 
      BackColor       =   &H0080FF80&
      Caption         =   "Скопировать файлы"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Копирует все MP3-файлы из списка в одну папку, чтобы вы могли их легко перенести на свой сервер"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdCreateScript 
      BackColor       =   &H0080FF80&
      Caption         =   "Создать конфиг"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Создаёт файл plugin.disco.cfg, который нужно поместить на сервер в /addons/sourcemod/configs/"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CheckBox chkDelete 
      BackColor       =   &H000000FF&
      Caption         =   "DEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Удалить из списка ..."
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0080FF80&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   20.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Добавить в список ..."
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   600
      Left            =   120
      Picture         =   "Form1.frx":948C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Сместить ниже"
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   600
      Left            =   120
      Picture         =   "Form1.frx":9A7E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Сместить выше"
      Top             =   1080
      Width           =   615
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3015
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777164
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "№"
         Object.Width           =   653
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Исполнитель"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Название"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Альбом"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Жанр"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Длина"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Размер"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Поток"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Год"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Путь к файлу"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.Label lblHost 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Обработчик:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   14
      Top             =   3800
      Width           =   990
   End
   Begin VB.Label lblПодсказка2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Подсказка 2. Двойной клик по пункту, чтобы воспроизвести трек."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   5145
   End
   Begin VB.Label lblПодсказка1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Поддерживаются только файлы MP3."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   9
      Top             =   3360
      Width           =   2910
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Загрузка ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   9240
      TabIndex        =   8
      Top             =   4080
      Width           =   990
   End
   Begin VB.Label lblTip2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Подсказка 3. Правая кнопка мыши по пункту, чтобы изменить его свойства."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   5895
   End
   Begin VB.Label lblTip1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Подсказка 1. Вы можете перетаскивать файлы (папки) на окно списка (из проводника или AIMP)."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   7515
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Добавление объекта"
      Begin VB.Menu mnuAddFile 
         Caption         =   "Добавить файл(ы) ..."
      End
      Begin VB.Menu mnuAddFolder 
         Caption         =   "Добавить из папки (*.mp3) ..."
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Контекстное меню"
      Begin VB.Menu mnuContextPlay 
         Caption         =   "Воспроизвести"
      End
      Begin VB.Menu mnuContextDelim1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextChangeAuthor 
         Caption         =   "Изменить автора"
      End
      Begin VB.Menu mnuContextChangeName 
         Caption         =   "Изменить название"
      End
      Begin VB.Menu mnuContextDelim2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextDelete 
         Caption         =   "Удалить из списка"
      End
      Begin VB.Menu mnuContextDeleteFile 
         Caption         =   "Удалить файл в корзину"
      End
      Begin VB.Menu mnuContextDelim3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextSelAll 
         Caption         =   "Выделить все строки"
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "Список"
      Begin VB.Menu mnuListCreate 
         Caption         =   "Создать новый"
      End
      Begin VB.Menu mnuListLoad 
         Caption         =   "Загрузить из файла..."
      End
      Begin VB.Menu mnuListSave 
         Caption         =   "Сохранить"
      End
      Begin VB.Menu mnuListSaveAs 
         Caption         =   "Сохранить как..."
      End
      Begin VB.Menu mnuListUseAsDefault 
         Caption         =   "Назначить открываемым по умолчанию"
      End
      Begin VB.Menu mnuListClearAll 
         Caption         =   "Очистить всё"
      End
   End
   Begin VB.Menu mhuHelp 
      Caption         =   "Справка"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "О программе"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const APP_NAME      As String = "DiscoList"
Const APP_TITLE     As String = APP_NAME & " Tool by Dragokas"
Const APP_VERSION   As String = "1.2"

Private Type tagINITCOMMONCONTROLSEX
    dwSize  As Long
    dwICC   As Long
End Type

Private Type ART_SETTINGS
    ScriptPath      As String  'путь к скрипту по умолчанию
    IniPath         As String  'путь к файлу настроек программы
    DefaultListPath As String  'путь к списку по умолчанию
    ListPath        As String  'путь к текущему открытому списку
    ListDelim       As String
    DefaultPauseSec As Long
    bSaved          As Boolean 'флаг, есть ли изменения
    RootFolder      As String  'корневая папка (точка отсчёта для относительного пути, если поставлена соответствующая галочка)
End Type

Private Type ART_NAMES
    MsgError        As String
    MsgDataInput    As String
End Type

Private Type ART_COLUMNS_ID
    Number      As Long
    Author      As Long
    Title       As Long
    Album       As Long
    Genre       As Long
    length      As Long
    Size        As Long
    Stream      As Long
    Year        As Long
    Path        As Long
End Type

Private Type ART_BASE
    Name        As ART_NAMES
    ColumnID    As ART_COLUMNS_ID
    Settings    As ART_SETTINGS
End Type

Private Art As ART_BASE

Private oShell As Object
Private bLockSort As Boolean

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function SetWindowTheme Lib "UxTheme.dll" (ByVal hwnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long

Const ICC_STANDARD_CLASSES As Long = &H4000&


Private Sub chkSort_Click()
    If chkSort.value = vbChecked Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
End Sub

Private Sub Form_Initialize()
    Dim hLib As Long, ICC As tagINITCOMMONCONTROLSEX
    hLib = LoadLibrary(StrPtr("shell32.dll"))
    With ICC
        .dwSize = Len(ICC)
        .dwICC = ICC_STANDARD_CLASSES
    End With
    Call InitCommonControlsEx(ICC)
    If hLib <> 0 Then FreeLibrary hLib
    Set oShell = CreateObject("Shell.Application")
End Sub

Private Sub Form_Load()
    mnuAdd.Visible = False
    mnuContext.Visible = False
    
    'Загрузка констант
    With Art
        With .Name
            .MsgError = APP_NAME & ": ошибка"
            .MsgDataInput = APP_NAME & ": ввод данных"
        End With
        With .ColumnID
            .Number = 0
            .Author = 1
            .Title = 2
            .Album = 3
            .Genre = 4
            .length = 5
            .Size = 6
            .Stream = 7
            .Year = 8
            .Path = 9
        End With
        With .Settings
            .DefaultListPath = "DiscoList_1.dlt"  'путь к файлу-списку автозапуска по умолчанию
            .ScriptPath = "musicconfig.cfg"
            .ListDelim = "*^_^*"                                      'разделитель полей в файле-списке
            .IniPath = BuildPath(App.Path, "main.ini")
        End With
    End With
    
    SaveControlsPos
    
    Dim cCtl As Control
    For Each cCtl In Me.Controls
        If TypeName(cCtl) = "CheckBox" Then
            SetWindowTheme cCtl.hwnd, StrPtr(" "), StrPtr(" ")
        End If
    Next
    
    'загрузка основных настроек программы
    LoadMainSettings
    
    'загрузка файла-списка автозагрузки по умолчанию
    If Art.Settings.ListPath <> "" Then
        LoadListFromFile Art.Settings.ListPath
        Call MarkSaved
    Else
        Call MarkUnsaved
    End If
    
    'галочка на пункте "Назначить открываемым по умолчанию"
    RefreshOpenAsDefaultCheckBox
    
    'надпись о состоянии работы прграммы
    lblState.Caption = ""
End Sub

Private Sub SaveControlsPos()
    'поддержка смещения контролов при изменении размеров формы (начальная позиция сохраняется в тег)
    Dim cCtl As Control
    For Each cCtl In Me.Controls
        If TypeName(cCtl) = "CommandButton" Or TypeName(cCtl) = "CheckBox" _
            Or TypeName(cCtl) = "Label" Or TypeName(cCtl) = "OptionButton" _
            Or TypeName(cCtl) = "TextBox" Then
                cCtl.Tag = CStr(cCtl.Top)
        End If
    Next
    Me.Tag = Me.Height
    lv.Tag = lv.Top + lv.Height
End Sub

Private Sub cmdDown_Click()
    Dim bPause As Boolean
    Dim nNewPos As Long
    Dim i As Long
    Dim nSelIdx As Long
    
    If lv.ListItems.Count = 0 Then Exit Sub
    
    'что перемещаем?
    'bPause = IsPause(lv.SelectedItem.SubItems(Art.ColumnID.Path))
    
    'поиск позиции для перемещения элемента
    For i = lv.SelectedItem.index + 1 To lv.ListItems.Count
        nNewPos = i
        Exit For
    Next
    If nNewPos = 0 Then Exit Sub 'если некуда перемещать
    
    'поменять местами элементы
    nSelIdx = lv.SelectedItem.index
    ListViewExchange nSelIdx, nNewPos
    'переместить признак "выделенный" элемент
    lv.ListItems(nSelIdx).Selected = False
    lv.ListItems(nNewPos).Selected = True
    
    lv.SetFocus
End Sub

Private Sub cmdCopyFiles_Click()
    
    Dim i As Long
    Dim sPath As String
    Dim sDestFolder As String
    Dim sNewFileName As String
    
    sDestFolder = BuildPath(App.Path, "music")

    If Not FolderExists(sDestFolder) Then MkDirW sDestFolder

    For i = 1 To lv.ListItems.Count
        sPath = lv.ListItems(i).SubItems(Art.ColumnID.Path)
        
        If FileExists(sPath) Then
            lblState.Caption = "Копирование: " & i & "/" & lv.ListItems.Count
            Me.Refresh
            sNewFileName = LCase$(Replace$(Translit(GetFileNameAndExt(sPath)), " ", "_"))
            'source != dest
            If StrComp(sPath, BuildPath(sDestFolder, sNewFileName), 1) <> 0 Then
                FileCopyW sPath, BuildPath(sDestFolder, sNewFileName), False
            End If
            DoEvents
        End If
    Next
End Sub

Private Sub cmdUp_Click()
    Dim bPause As Boolean
    Dim nNewPos As Long
    Dim i As Long
    Dim nSelIdx As Long
    
    If lv.ListItems.Count = 0 Then Exit Sub
    
    'поиск позиции для перемещения элемента
    For i = lv.SelectedItem.index - 1 To 1 Step -1
        nNewPos = i
        Exit For
    Next
    If nNewPos = 0 Then Exit Sub 'если некуда перемещать
    
    'поменять местами элементы
    nSelIdx = lv.SelectedItem.index
    ListViewExchange nSelIdx, nNewPos
    'переместить признак "выделенный" элемент
    lv.ListItems(nSelIdx).Selected = False
    lv.ListItems(nNewPos).Selected = True
    
    lv.SetFocus
End Sub

'обмен элементов местами
Private Sub ListViewExchange(nIdx1 As Long, nIdx2 As Long)
    Dim i&, cntSubItems&
    cntSubItems = lv.ListItems(nIdx1).ListSubItems.Count
    ReDim Rec_1(1 To cntSubItems) As String
    ReDim Rec_2(1 To cntSubItems) As String
    ReDim Rec_1_Color(1 To cntSubItems) As Long
    ReDim Rec_2_Color(1 To cntSubItems) As Long
    
    'сохраняем первый элемент
    For i = 1 To UBound(Rec_1)
        Rec_1(i) = lv.ListItems(nIdx1).SubItems(i)
        Rec_1_Color(i) = lv.ListItems(nIdx1).ListSubItems(i).ForeColor
    Next
    'сохраняем второй элемент
    For i = 1 To UBound(Rec_2)
        Rec_2(i) = lv.ListItems(nIdx2).SubItems(i)
        Rec_2_Color(i) = lv.ListItems(nIdx2).ListSubItems(i).ForeColor
    Next
    'обмен
    For i = 1 To UBound(Rec_1)
        lv.ListItems(nIdx2).SubItems(i) = Rec_1(i)
        lv.ListItems(nIdx2).ListSubItems(i).ForeColor = Rec_1_Color(i)
    Next
    For i = 1 To UBound(Rec_2)
        lv.ListItems(nIdx1).SubItems(i) = Rec_2(i)
        lv.ListItems(nIdx1).ListSubItems(i).ForeColor = Rec_2_Color(i)
    Next
    
    MarkUnsaved
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Art.Settings.bSaved = False Then
        If MsgSaveCancelled() Then Cancel = True
    End If
    SaveMainSettings
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Static isInit As Boolean
    
    If Not isInit Then 'чтобы дать время форме прогрузить все контролы
        isInit = True
        Exit Sub
    End If
    
    Dim lInitFormHeight As Long
    Dim lInitBottomLvCoord As Long
    Dim lMinFormHeight As Long
    
    lInitFormHeight = (CLng(Val(Me.Tag)))
    lInitBottomLvCoord = (CLng(Val(lv.Tag)))
    
    lMinFormHeight = lInitFormHeight
    
    'ограничение минимального размера окна
    If frmMain.Width < 3150 Then frmMain.Width = 3150
    If frmMain.Height < lMinFormHeight Then frmMain.Height = lMinFormHeight
    
    If frmMain.Width >= 3120 Then lv.Width = frmMain.Width - 1245
    If frmMain.Height >= 5295 Then lv.Height = frmMain.Height - (lInitFormHeight - lInitBottomLvCoord) - lv.Top  '- 2280
    
    Dim cCtl As Control
    Dim lInitControlTop As Long
    
    'поддержка смещения контролов, расположенных ниже ListView, при изменении размеров формы
    For Each cCtl In Me.Controls
        If TypeName(cCtl) = "CommandButton" Or TypeName(cCtl) = "CheckBox" Or _
            TypeName(cCtl) = "Label" Or TypeName(cCtl) = "OptionButton" Or _
            TypeName(cCtl) = "TextBox" Then
            
            lInitControlTop = CLng(Val(cCtl.Tag))
            If lInitControlTop > lInitBottomLvCoord Then
                cCtl.Top = Me.Height - (lInitFormHeight - lInitControlTop) 'current height minus delta
            End If
        End If
    Next
    
    Me.Refresh
End Sub

Private Sub PlayFile(sFile As String)
    oShell.ShellExecute sFile, "", "", "", 7 '7 - no active, minimized
End Sub

Private Sub PlaySelectedItem()
    Dim sFile As String
    If LVSelCount <> 0 Then
        sFile = lv.SelectedItem.SubItems(Art.ColumnID.Path)
        If FileExists(sFile) Then
            PlayFile sFile
        End If
    End If
End Sub

' вызов контекстного меню по умолчанию при двойном клике по пункту
Private Sub lv_DblClick()
    PlaySelectedItem     'воспроизвести трек
End Sub

'удалить выделенныый эл-т из списка (кнопкой на клавиатуре: Delete или BackSpace)
Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
      Case 46, 8
        If lv.ListItems.Count <> 0 Then
            With lv.SelectedItem
                'If .Selected Then
                lv.ListItems.Remove .index ' удаляю из списка
            End With
        End If
      Case 13
        lv_DblClick
    End Select
End Sub

'кнопка "+" (добавить объект)
Private Sub cmdAdd_Click()
    PopupMenu mnuAdd
End Sub

'контекстное меню
Private Sub lv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim n&
    If Button = 2 Then
        With lv.ListItems
            For n = 1 To .Count
                If .Item(n).Selected Then
                    PopupMenu mnuContext
                    Exit Sub
                End If
            Next
        End With
    End If
End Sub

'Меню: добавить файл...
Private Sub mnuAddFile_Click()
    Dim sFile As String
    Dim aFiles() As String
    Dim i As Long
    
    sFile = GetOpenFile(Me.hwnd, True)
    
    If sFile <> "" Then
        aFiles = Split(sFile, vbNullChar)
        
        For i = 0 To UBound(aFiles)
            AddFile aFiles(i)
        Next
        
        Me.Refresh
        'MarkUnsaved
    End If
End Sub

'Меню: добавить папку ...
Private Sub mnuAddFolder_Click()
    Dim sFolder As String
    sFolder = OpenFolderDialog(Me.hwnd, App.Path)
    
    If sFolder <> "" Then
        lblState.Caption = "Загрузка ..."
        lv.Enabled = False
        AddFolder sFolder
        'MarkUnsaved
        lv.Enabled = True
        lblState.Caption = ""
    End If
End Sub

Private Sub LvSort()
    On Error GoTo ErrorHandler:
    'Sorting user type array using bufer array of positions (c) Dragokas
    Dim i As Long, j As Long
    Dim cnt As Long
    Dim aColumns() As String
    Dim names() As String
    
    If chkSort.value = vbUnchecked Then Exit Sub
    
    cnt = lv.ListItems.Count
    
    ReDim names(cnt) As String
    
    For i = 1 To cnt
        With lv.ListItems(i)
            'key of sort is Author + Title + delim + all.lines
            names(i) = .ListSubItems(Art.ColumnID.Author).Text & _
                " - " & .ListSubItems(Art.ColumnID.Title).Text & Art.Settings.ListDelim & LVConcatLine(i)
        End With
    Next
    
    QuickSort names, 1, cnt
    
    lv.ListItems.Clear
    
    For i = 1 To cnt
        aColumns = Split(names(i), Art.Settings.ListDelim)
        
        With lv.ListItems.Add(i, , i)
            For j = 2 To UBound(aColumns)
                .SubItems(j - 1) = aColumns(j)
            Next
            
            'выделить красным, если путь в юникоде
            If Not FileExists(.SubItems(Art.ColumnID.Path)) Then
                '.ListSubItems(Art.ColumnID.Path).ForeColor = &HFF&
                '.ForeColor = &HFF&
                LvMarkRed i
            End If
        End With
    Next
    Exit Sub
ErrorHandler:
    Debug.Print "Error in LvSort. Err # " & Err.Number & ": " & Err.Description
End Sub

Private Sub AddFolder(sFolder As String)
    Dim i As Long
    Dim aFiles() As String
    
    aFiles = ListFiles(sFolder, ".mp3", True)
    
    bLockSort = True
    If IsArrDimmed(aFiles) Then
        For i = 0 To UBound(aFiles)
            AddFile aFiles(i)
            If i Mod 20 = 0 Then DoEvents
        Next
    End If
    bLockSort = False
    
    Call LvSort
    
    Me.Refresh
End Sub

Private Sub AddFile(sFile As String)
    On Error GoTo ErrorHandler:

    Dim nPos As Long
    Dim Artist As String
    Dim Title As String
    Dim Genre As String
    Dim Album As String
    Dim i As Long

    If StrComp(GetExtensionName(sFile), ".mp3", 1) <> 0 Then Exit Sub
    
    MarkUnsaved
    
    Dim oDir As Object:     Set oDir = oShell.Namespace(GetParentDir(sFile))
    Dim oFile As Object:    Set oFile = oDir.ParseName(GetFileNameAndExt(sFile))
    
    If lv.ListItems.Count = 0 Then
        nPos = 1
    Else
        nPos = lv.SelectedItem.index + 1
    End If
    
        '№
    With lv.ListItems.Add(nPos, , 0)
        'author
        Artist = Trim(oDir.getdetailsof(oFile, 20))
        Title = Trim(oDir.getdetailsof(oFile, 21))
        Genre = oDir.getdetailsof(oFile, 16)
        Album = oDir.getdetailsof(oFile, 14)
        
        If Artist = "artist" Then Artist = ""
        If Title = "title" Then Title = ""
        If Genre = "genre" Then Genre = ""
        If Album = "title" Then Album = ""
        
        .SubItems(Art.ColumnID.Author) = IIf(Artist = "", GetFileName(sFile), Artist)
        'title
        .SubItems(Art.ColumnID.Title) = Title
        'Album
        .SubItems(Art.ColumnID.Album) = Album
        'genre
        .SubItems(Art.ColumnID.Genre) = Genre
        'length
        .SubItems(Art.ColumnID.length) = oDir.getdetailsof(oFile, 27)
        'size
        .SubItems(Art.ColumnID.Size) = oDir.getdetailsof(oFile, 1)
        'stream
        .SubItems(Art.ColumnID.Stream) = FormatStrStream(oDir.getdetailsof(oFile, 28))
        'year
        .SubItems(Art.ColumnID.Year) = oDir.getdetailsof(oFile, 15)
        'path
        .SubItems(Art.ColumnID.Path) = sFile
        
        'ANSI-to-Unicode mind hell
        Artist = .SubItems(Art.ColumnID.Author)
        Title = .SubItems(Art.ColumnID.Title)
        
        Artist = Trim(Replace$(Artist, "?", ""))
        Title = Trim(Replace$(Title, "?", ""))
        
        .SubItems(Art.ColumnID.Title) = Title
        .SubItems(Art.ColumnID.Author) = IIf(Artist = "", GetFileName(sFile), Artist)
        
        Artist = .SubItems(Art.ColumnID.Author)
        Title = .SubItems(Art.ColumnID.Title)
        
        Artist = Trim(Replace$(Artist, "?", ""))
        If Artist = "" Then Artist = "Track_" & nPos
        Title = Trim(Replace$(Title, "?", ""))
        
        'Title = Translit(Title)
        'Artist = Translit(Artist)
        
        .SubItems(Art.ColumnID.Title) = Title
        .SubItems(Art.ColumnID.Author) = Artist
        
        'выделить красным, если путь в юникоде
        If Not FileExists(.SubItems(Art.ColumnID.Path)) Then
            '.ListSubItems(Art.ColumnID.Path).ForeColor = &HFF&
            '.ForeColor = &HFF&
            LvMarkRed nPos
        End If
    End With
    
    Dim Times As Long
    'проверка на дубликат
    For i = 1 To lv.ListItems.Count
        With lv.ListItems(i)
            If Artist = .ListSubItems(Art.ColumnID.Author).Text And _
              Title = .ListSubItems(Art.ColumnID.Title).Text Then
            
                Times = Times + 1
                If Times > 1 Then
                    lv.ListItems.Remove (nPos)
                    Exit For
                End If
            End If
        End With
    Next
    
    'lv.SelectedItem.Selected = False
    'lv.ListItems(nPos).Selected = True
    
    'перенумеровать пункты
    'Call ReEnum
    
    If Not bLockSort Then Call LvSort
    
    Set oFile = Nothing
    Set oDir = Nothing
    
    Exit Sub
ErrorHandler:
    Debug.Print "AddFile. Error: " & Err.Number & " - " & Err.Description & ". File: " & sFile
End Sub

Private Sub LvRemoveDuplicates()
    Dim i As Long
    Dim j As Long
    
    Dim Artist As String
    Dim Title As String

    For i = lv.ListItems.Count To 1 Step -1
        Artist = lv.ListItems(i).ListSubItems(Art.ColumnID.Author).Text
        Title = lv.ListItems(i).ListSubItems(Art.ColumnID.Title).Text
    
        For j = i - 1 To 1 Step -1
    
            If Artist = lv.ListItems(j).ListSubItems(Art.ColumnID.Author).Text And _
              Title = lv.ListItems(j).ListSubItems(Art.ColumnID.Title).Text Then
                lv.ListItems.Remove (j)
            End If
        Next
    Next
End Sub

Private Function FormatStrStream(s$) As String
    On Error GoTo ErrorHandler:
    FormatStrStream = Val(Replace(s, ChrW(8206), ""))
    Exit Function
ErrorHandler:
    Debug.Print "AddFile. Error: " & Err.Number & " - " & Err.Description
End Function

Private Sub mnuAutorunAddScriptToAutorun_Click()
    MsgBox "Пока не реализовано"
End Sub

'Контекстное: изменить название
Private Sub mnuContextChangeName_Click()
    Dim n&, sNewName$, sOldName$
    If LVSelCount > 1 Then
        MsgBox "Выделите не больше 1 элемента!", vbExclamation, Art.Name.MsgError
        Exit Sub
    End If
    With lv.SelectedItem
        sOldName = .SubItems(Art.ColumnID.Title)
        sNewName = InputBox("Введите новое название:", Art.Name.MsgDataInput, sOldName)
        If StrPtr(sNewName) <> 0 Then
            .SubItems(Art.ColumnID.Title) = sNewName
            MarkUnsaved
        End If
    End With
End Sub

'Контекстное: изменить путь
Private Sub mnuContextChangeAuthor_Click()
    Dim n&, sNewName$, sOldName$
    If LVSelCount > 1 Then
        MsgBox "Выделите не больше 1 элемента!", vbExclamation, Art.Name.MsgError
        Exit Sub
    End If
    With lv.SelectedItem
        sOldName = .SubItems(Art.ColumnID.Author)
        sNewName = InputBox("Введите новое название:", Art.Name.MsgDataInput, sOldName)
        If StrPtr(sNewName) <> 0 Then
            .SubItems(Art.ColumnID.Author) = sNewName
            MarkUnsaved
        End If
    End With
End Sub

'Контекстное: удалить пункт из списка
Private Sub mnuContextDelete_Click()
    Call DeleteSelItem(False)
End Sub

'Удалить пункт
Private Sub chkDelete_Click()
    Static isClicked As Boolean
    If Not isClicked Then
        isClicked = True
        Call DeleteSelItem(False)
        chkDelete.value = 0
    Else
        isClicked = False
    End If
End Sub

Private Sub DeleteSelItem(bToRecycleBin As Boolean)
    Dim nSelIdx As Long
    Dim cntSel As Long
    Dim sPath As String
    Dim i As Long
    
    With lv.ListItems
        If .Count <> 0 Then
            nSelIdx = lv.SelectedItem.index
            'перечислить все выделенные элементы
            For i = .Count To 1 Step -1
                If lv.ListItems(i).Selected Then
                    If bToRecycleBin Then
                        sPath = lv.ListItems(i).ListSubItems(Art.ColumnID.Path).Text
                        If FileExists(sPath) Then
                            SendFileToRecycleBin sPath, False, True
                        End If
                    End If
                    cntSel = cntSel + 1
                    .Remove i
                End If
            Next
            'выделить предыдущий пункт, если был выделен не более чем 1
            If lv.ListItems.Count <> 0 And cntSel = 1 Then
                If nSelIdx - 1 <> 0 Then
                    lv.ListItems(nSelIdx - 1).Selected = True
                Else
                    lv.ListItems(nSelIdx).Selected = True
                End If
                
                lv.SetFocus
            End If
            'перенумеровать пункты
            Call ReEnum
            MarkUnsaved
        End If
    End With
    
    MarkUnsaved
End Sub

'перенумеровать пункты
Private Sub ReEnum()
    Dim nItem As Long
    Dim i As Long
    With lv.ListItems
        If .Count <> 0 Then
            For i = 1 To .Count
                'If Not IsPause(, i) Then
                    nItem = nItem + 1
                    lv.ListItems(i).Text = nItem
                'End If
            Next
        End If
    End With
End Sub

'Контекстное: удалить файл в корзину
Private Sub mnuContextDeleteFile_Click()
    Call DeleteSelItem(True)
End Sub

'Контекстное: воспроизвести
Private Sub mnuContextPlay_Click()
    PlaySelectedItem
End Sub

'Контекстное: выделить всё
Private Sub mnuContextSelAll_Click()
    Dim n&
    With lv.ListItems
        For n = 1 To .Count
            'If Not IsPause(.Item(n).SubItems(Art.ColumnID.Path)) Then
                If Not .Item(n).Selected Then
                    .Item(n).Selected = True
                End If
'            Else
'                If .Item(n).Selected Then
'                    .Item(n).Selected = False
'                End If
'            End If
        Next
    End With
End Sub

'поддержка Drag & Drop
Private Sub AddObjToList(Data As Object)
    Const vbCFFiles As Long = 15&
    Dim Obj
    If Data.GetFormat(vbCFFiles) Then
        For Each Obj In Data.Files
            If FolderExists(CStr(Obj)) Then
                AddFolder CStr(Obj)
            ElseIf FileExists(CStr(Obj)) Then
                AddFile CStr(Obj)
            End If
        Next
    End If
End Sub

Private Sub lv_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    AddObjToList Data
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    AddObjToList Data
End Sub

'выделяет красным файлы в списке, которые не существуют
Private Sub CheckFilesPresence()
    Dim i As Long
    Dim sFile As String
    For i = 1 To lv.ListItems.Count
        With lv.ListItems(i).ListSubItems(Art.ColumnID.Path)
            'If Not IsPause(.Text) Then
                sFile = .Text
                If Art.Settings.RootFolder <> "" Then sFile = BuildPath(Art.Settings.RootFolder, Mid$(sFile, 3))
                If Not FileExists(sFile) Then
                    lv.ListItems(i).ForeColor = &HFF&
                End If
            'End If
        End With
    Next
End Sub

'возвращает количество выделенных элементов
Private Function LVSelCount() As Long
    Dim i As Long
    For i = 1 To lv.ListItems.Count
        If lv.ListItems(i).Selected Then LVSelCount = LVSelCount + 1
    Next
End Function

'Файл => выход
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

'Список => Очистить всё
Private Sub mnuListClearAll_Click()
    If MsgBox("Вы уверены, что хотите очистить весь список?", vbQuestion Or vbYesNo, APP_NAME) = vbYes Then
        lv.ListItems.Clear
        MarkUnsaved
    End If
End Sub

'Справка => О программе
Private Sub mnuHelpAbout_Click()
    MsgBox APP_NAME & vbTab & "(версия " & APP_VERSION & ")" & vbCrLf & vbCrLf & _
        "Программа создания файла конфигурации для плагина SourceMod 'Disco Mod' от MitchDizzle_." & vbCrLf & vbCrLf & _
        "Автор: Alex Dragokas" & vbCrLf & vbCrLf & _
        "Лицензия: GNU GPLv3", vbInformation, ""
End Sub

'Контекстное: Настроить как ярлык ...
Private Sub mnuContextSetShortcut_Click()
    MsgBox "Пока не реализовано"
End Sub

'Список => Создать новый
Private Sub mnuListCreate_Click()
    If Art.Settings.bSaved = False Then
        If MsgSaveCancelled() Then Exit Sub
    End If
    lv.ListItems.Clear
    Art.Settings.ListPath = ""
    MarkSaved
    RefreshOpenAsDefaultCheckBox
End Sub

'Обновить состояние галочки для пункта "Назначить открываемым по умолчанию"
Private Sub RefreshOpenAsDefaultCheckBox()
    If Art.Settings.ListPath = Art.Settings.DefaultListPath Then
        mnuListUseAsDefault.Checked = True
    Else
        mnuListUseAsDefault.Checked = False
    End If
End Sub

'Диалог запроса на сохранение списка
Private Function MsgSaveCancelled() As Boolean
    'returns true if user clicked "Cancel"
    
    Select Case MsgBox("Список был изменён. Сохранить?", vbQuestion Or vbYesNoCancel, APP_NAME)
    Case vbCancel
        MsgSaveCancelled = True
    Case vbYes
        If Art.Settings.ListPath = "" Then 'call Save As...
            If ListSaveAsCancelled() Then MsgSaveCancelled = True
        Else
            mnuListSave_Click
        End If
    End Select
End Function

'Список => Загрузить из файла ...
Private Sub mnuListLoad_Click()
    Dim sFile As String
    
    'проверка, сохранён ли текущий документ
    If Art.Settings.bSaved = False Then
        If MsgSaveCancelled() Then Exit Sub
    End If
    
    sFile = GetOpenFile(Me.hwnd, , "Файлы конфигурации" & vbNullChar & "*.dlt" & vbNullChar)
    
    If sFile <> "" Then
        lblState.Caption = "Загрузка ..."
        lv.Enabled = False
        LoadListFromFile sFile
        lv.Enabled = True
        lblState.Caption = ""
    End If
    
    RefreshOpenAsDefaultCheckBox
End Sub

'загрузка основных настроек программы
Private Sub LoadMainSettings()
    ' чтение пути к списку автозапуска по умолчанию
    If FileExists(Art.Settings.IniPath) Then
        Art.Settings.ListPath = ReadIniValue(Art.Settings.IniPath, "General", "DefaultListPath", Art.Settings.DefaultListPath)
        'full path -> relative
        If StrBeginWith(Art.Settings.ListPath, App.Path & "\") Then
            Art.Settings.ListPath = Mid(Art.Settings.ListPath, Len(App.Path & "\") + 1)
        End If
        Art.Settings.DefaultListPath = Art.Settings.ListPath
    Else
        Art.Settings.ListPath = Art.Settings.DefaultListPath
    End If
End Sub

'сохранение основных настроек программы
Private Sub SaveMainSettings()
    WriteIniValue Art.Settings.IniPath, "General", "DefaultListPath", Art.Settings.DefaultListPath
End Sub

'загрузка настроек из указанного файла
Private Sub LoadListFromFile(sFile As String)
    On Error GoTo ErrorHandler
    Dim nItems As Long
    Dim i As Long
    Dim j As Long
    Dim sTmp As String
    Dim aColumns() As String
    
    'relative path -> full
    If Mid$(sFile, 2, 1) <> ":" Then sFile = BuildPath(App.Path, sFile)
    
    If Not FileExists(sFile) Then Exit Sub
    
    lv.ListItems.Clear
    
    sTmp = ReadIniValue(sFile, "General", "Count", "Failure")
    
    If sTmp = "Failure" Or sTmp = "" Then
        MsgBox "Ошибка: выбранный список повреждён!", vbCritical, Art.Name.MsgError
        Exit Sub
    Else
        nItems = Val(sTmp)
    End If
    
    sTmp = ReadIniValue(sFile, "General", "Host", "Failure")
    If sTmp <> "Failure" And sTmp <> "" Then
        txtHost.Text = sTmp
    End If
    
    For i = 1 To nItems
        aColumns = Split(ReadIniValue(sFile, "List", "Item_" & i), Art.Settings.ListDelim)
    
        With lv.ListItems.Add(i, , aColumns(0))
            For j = 1 To UBound(aColumns)
                If j = Art.ColumnID.Author Or j = Art.ColumnID.Title Then
                    .SubItems(j) = aColumns(j)
                Else
                    .SubItems(j) = aColumns(j)
                End If
            Next
        End With
        
        If i Mod 20 = 0 Then DoEvents
    Next
    
    LvRemoveDuplicates
    LvSort
    
    'выделяет красным файлы в списке, которые не существуют
    Call CheckFilesPresence
    
    lv.Refresh
    Art.Settings.ListPath = sFile
    MarkSaved
    
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка: некорректный файл конфигурации!", vbCritical, Art.Name.MsgError
    Resume Next
End Sub

'Список => Сохранить
Private Sub mnuListSave_Click()
    Dim nItems As Long
    Dim i As Long
    
    lblState.Caption = "Сохранение ..."
    
    nItems = lv.ListItems.Count
    
    WriteIniValue Art.Settings.ListPath, "General", "Count", nItems
    WriteIniValue Art.Settings.ListPath, "General", "Host", txtHost.Text

    For i = 1 To nItems
        WriteIniValue Art.Settings.ListPath, "List", "Item_" & i, LVConcatLine(i)
        If i Mod 20 = 0 Then DoEvents
    Next
    
    MarkSaved
    lblState.Caption = ""
End Sub

'взять строку ListView, объединив все колонки через разделитель
Private Function LVConcatLine(nItem As Long) As String
    Dim i As Long
    Dim sStr As String
    
    sStr = lv.ListItems(nItem).Text
    
    For i = 1 To lv.ListItems(nItem).ListSubItems.Count
        sStr = sStr & Art.Settings.ListDelim & lv.ListItems(nItem).SubItems(i)
    Next
    
    LVConcatLine = sStr
End Function

'Список => Сохранить как...
Private Sub mnuListSaveAs_Click()
    Call ListSaveAsCancelled
End Sub

Private Function ListSaveAsCancelled() As Boolean
    'returns true if user clicked "Cancel"
    Dim sFile As String
    
    sFile = GetSaveFile(Me.hwnd)
    
    'full path -> relative
    If StrBeginWith(sFile, App.Path & "\") Then
        sFile = Mid(sFile, Len(App.Path & "\") + 1)
    End If
    
    If sFile = "" Then
        ListSaveAsCancelled = True
    Else
        Art.Settings.ListPath = sFile & IIf(UCase(GetExtensionName(sFile)) = ".DLT", "", ".dlt")
        mnuListSave_Click
    End If
    RefreshOpenAsDefaultCheckBox
End Function

'Список => Использовать как основной
Private Sub mnuListUseAsDefault_Click()
    If mnuListUseAsDefault.Checked Then
        Art.Settings.DefaultListPath = ""
    Else
        Art.Settings.DefaultListPath = Art.Settings.ListPath
    End If
    RefreshOpenAsDefaultCheckBox
End Sub

'Создать батник для выполнения последовательности автозапуска
Private Sub mnuAutorunCreateScript_Click()
    cmdCreateScript_Click
End Sub

Private Sub cmdCreateScript_Click()
    
    Dim i As Long
    Dim s As String
    Dim hFile As Long
    Dim sFullTitle As String
    Dim sTitle As String
    Dim sAuthor As String
    Dim sMusicFile As String
    Dim sHost As String
    Dim bAtLeast1File As Boolean
    
    'sHost = "dragokas.com/music/index.php?id="
    sHost = txtHost.Text
    
    s = """MusicConfig"""
    s = s & vbCrLf & "{"
    
    For i = 1 To lv.ListItems.Count
        sAuthor = lv.ListItems(i).SubItems(Art.ColumnID.Author)
        sTitle = lv.ListItems(i).SubItems(Art.ColumnID.Title)
        If sAuthor = "" Then
            sFullTitle = sTitle
        Else
            sFullTitle = sAuthor & IIf(sTitle <> "", " - " & sTitle, "")
        End If
        sMusicFile = lv.ListItems(i).SubItems(Art.ColumnID.Path)
        
        'If FileExists(sMusicFile) Then
            bAtLeast1File = True
            s = s & vbCrLf & vbTab & """" & sFullTitle & """"
            s = s & vbCrLf & vbTab & "{"
            s = s & vbCrLf & vbTab & vbTab & """" & "path" & """" & String$(4, vbTab) & _
                """" & sHost & URLEncode(Replace(Translit(GetFileName(sMusicFile)), " ", "_")) & GetExtensionName(sMusicFile) & """"
            s = s & vbCrLf & vbTab & "}"
        'End If
    Next
    
    If Not bAtLeast1File Then
        MsgBox "Ошибка: не найдено ни одного файла из списка!"
        Exit Sub
    End If
    
    s = s & vbCrLf & "}"
    
    If OpenW(Art.Settings.ScriptPath, FOR_OVERWRITE_CREATE, hFile) Then
        s = StrConv(s, vbFromUnicode)
        PutW hFile, 1, StrPtr(s), LenB(s)
        CloseW hFile
        'открываем папку со скриптом
        OpenFolderAndSelectItem Art.Settings.ScriptPath
    Else
        MsgBox "Ошибка: не могу открыть файл для записи: " & Art.Settings.ScriptPath, vbCritical, Art.Name.MsgError
    End If
End Sub

'добавить в заголовок окна пометку * (не сохранён)
Private Sub MarkUnsaved()
    If Art.Settings.bSaved Then
        Me.Caption = "*[" & GetFileNameAndExt(Art.Settings.ListPath) & "] - " & APP_TITLE
        Art.Settings.bSaved = False
    End If
End Sub

'[Пометить сохранённым] - убрать из заголовка окна пометку * (не сохранён)
Private Sub MarkSaved()
    Me.Caption = "[" & GetFileNameAndExt(Art.Settings.ListPath) & "] - " & APP_TITLE
    Art.Settings.bSaved = True
End Sub

Public Sub QuickSort(j() As String, ByVal low As Long, ByVal high As Long)
    On Error GoTo ErrorHandler:
    Dim i As Long, l As Long, M As String, wsp As String
    i = low: l = high: M = j((i + l) \ 2)
    Do Until i > l: Do While j(i) < M: i = i + 1: Loop: Do While j(l) > M: l = l - 1: Loop
        If (i <= l) Then wsp = j(i): j(i) = j(l): j(l) = wsp: i = i + 1: l = l - 1
    Loop
    If low < l Then QuickSort j, low, l
    If i < high Then QuickSort j, i, high
    Exit Sub
ErrorHandler:
    Debug.Print "Error in QuickSort. Err # " & Err.Number & ": " & Err.Description
End Sub

Private Sub LvMarkRed(nItem As Long)
    Dim i As Long
    For i = 1 To lv.ListItems(nItem).ListSubItems.Count
        lv.ListItems(nItem).ListSubItems(i).ForeColor = &HFF&
        lv.ListItems(nItem).ForeColor = &HFF&
    Next
End Sub

Private Sub txtHost_Change()
    MarkUnsaved
End Sub

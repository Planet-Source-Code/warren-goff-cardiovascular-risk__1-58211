VERSION 5.00
Begin VB.Form Cardiovascular 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cardiovascular Risk & Treatment Dyslipidemia"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   Icon            =   "Cardiovascular.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4320
      TabIndex        =   67
      Top             =   5040
      Width           =   5655
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Typical"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   73
         Top             =   0
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Atypical"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   72
         Top             =   0
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Non-"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   71
         Top             =   0
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Unknown"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   70
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "On Estrogen"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   69
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "No Estrogen"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   68
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Pain Symptoms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   76
         ToolTipText     =   "If chest pain is a complaint is it:"
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Anginal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   75
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Estrogen Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   74
         Top             =   360
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   5520
         Y1              =   285
         Y2              =   285
      End
   End
   Begin VB.TextBox CAD 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   24
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Interpret"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7200
      Width           =   975
   End
   Begin VB.Frame Conversions 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Conversions"
      Height          =   1095
      Left            =   7560
      TabIndex        =   65
      Top             =   3600
      Width           =   3255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2115
         ScaleHeight     =   315
         ScaleWidth      =   1035
         TabIndex        =   77
         Top             =   705
         Width           =   1035
         Begin VB.TextBox Outy 
            Height          =   285
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.TextBox Inny 
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Text            =   "Select Parameter"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   2
         X1              =   1920
         X2              =   2040
         Y1              =   960
         Y2              =   840
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   1
         X1              =   1920
         X2              =   2040
         Y1              =   720
         Y2              =   840
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   0
         X1              =   1320
         X2              =   2040
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.TextBox Interpret 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   5760
      Width           =   6975
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   23
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   22
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox diastolic 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   21
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox LDL 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      TabIndex        =   16
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   20
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   19
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   18
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2400
      TabIndex        =   17
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox Triglyceride 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      TabIndex        =   15
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Fibrinogen 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   9720
      TabIndex        =   31
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   9720
      TabIndex        =   30
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Obesity 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Height          =   285
      Left            =   9720
      TabIndex        =   29
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox FhxCAD 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Height          =   285
      Left            =   9720
      TabIndex        =   28
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Hyperlipidemia 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Height          =   285
      Left            =   9720
      TabIndex        =   27
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Hypertension 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Height          =   285
      Left            =   9720
      TabIndex        =   26
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   13
      Left            =   9720
      TabIndex        =   25
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   12
      Left            =   2400
      TabIndex        =   14
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   11
      Left            =   2400
      TabIndex        =   13
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   10
      Left            =   2400
      TabIndex        =   12
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   9
      Left            =   2400
      TabIndex        =   11
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   8
      Left            =   2400
      TabIndex        =   10
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   7
      Left            =   2400
      TabIndex        =   9
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Syst1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Has Coronary Artery Disease?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   3720
      TabIndex        =   66
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Does the patient have a first degree female relative who developed CHD before age 65? (Y or N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   28
      Left            =   3720
      TabIndex        =   64
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Does the patient have a first degree male relative who developed CHD before age 55? (Y or N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   27
      Left            =   3720
      TabIndex        =   63
      ToolTipText     =   "CHD is coronary heart disease."
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Diastolic BP :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   240
      TabIndex        =   62
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FHx of definite MI or sudden death before 65 father/other 1st degree male relative? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   25
      Left            =   3720
      TabIndex        =   61
      ToolTipText     =   $"Cardiovascular.frx":08CA
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LDL cholesterol in mg/dL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   240
      TabIndex        =   60
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FHx of definite MI or sudden death before 65 mother/other 1st degree female relative? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   3720
      TabIndex        =   59
      ToolTipText     =   $"Cardiovascular.frx":0961
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Is the person's blood pressure 140/90 mm Hg or greater on several occasions? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   22
      Left            =   3720
      TabIndex        =   58
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Was menopause premature and not tx'ed with estrogent? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   57
      ToolTipText     =   "Did menopause happen early and was it NOT treated with estrogen or hormone replacement therapy (HRT)?"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Postmenopausal? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   56
      ToolTipText     =   "Have you passed Menopause?"
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Triglycerides in mg/dL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   55
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plasma fibrinogen in mg/dL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   20
      Left            =   7560
      TabIndex        =   54
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Body Mass Index:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   7560
      TabIndex        =   53
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Obesity? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   7560
      TabIndex        =   52
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FHx of CAD? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   7560
      TabIndex        =   51
      ToolTipText     =   "Is there a Family History of coronary artery disease?"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "History Hyperlipidemia? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   7560
      TabIndex        =   50
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "History Hypertension? (Y\N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   7560
      TabIndex        =   49
      ToolTipText     =   "Do you have a history of high blood pressure?"
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Symptoms poss CAD? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   240
      TabIndex        =   48
      ToolTipText     =   "Sypmtoms of possible Coronary Artery Disease (Disease of the arteries around the heart)"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Left Ventricular Hypertrophy?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   7560
      TabIndex        =   47
      ToolTipText     =   "Is there thickening of the muscle of the left side of the heart?"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current HDL cholesterol level, in mg/dL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   240
      TabIndex        =   46
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Total Cholesterol level, in mg/dL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   240
      TabIndex        =   45
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current BP meds? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   44
      ToolTipText     =   "Do you currently take blood pressure (hypertension) pills (medications)?"
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Diabetic? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   43
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Do you smoke? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   42
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rarely exercise or do anything physical? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   240
      TabIndex        =   41
      ToolTipText     =   "Do you live a sedentary life style and rarely exert yourself?"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FHx of Heart ds or MI prior to age of 60? (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   240
      TabIndex        =   40
      ToolTipText     =   "Family History of heart diseas or Heart attack prior to the age of 60?"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Systolic BP :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   39
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight (pounds):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   38
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Height (inches):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   37
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Age (years):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   36
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender (M/F):: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Cardiovascular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AHA As Integer
Dim AHArisk As String
Dim BMIcaption As String
Dim BmiValue As Double

Private Sub Bmi()
On Error Resume Next

BmiValue = (Int(((Val(Text1(3).Text) / ((Val(Text1(2).Text) ^ 2)) * 70300) / 100) * 100)) / 100


If Text1(0) = "M" Then
Select Case BmiValue

Case Is < 20.69
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are UNDERWEIGHT. The lower the BMI the greater the risk of Malnutrition.   You should supplement each of your THREE meals daily with a Formula 1 Smoothie until your BMI is normal (between 20.7 and 26.4).   Then go on maintenance."
Obesity = "N"

Case 20.7 To 26.4
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are at a NORMAL WEIGHT.  Your risk of Co-morbidities is very low.   You should substitute ONE of your usual THREE meals daily with a ONE Formula 1 Smoothie and maintain your BMI as is (between 20.7 and 26.4).   Congratulations!   Stay on maintenance."
Obesity = "N"

Case 26.41 To 27.8
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are MARGINALLY OVERWEIGHT.  Some risk of Obesity Co-morbidities exists but they are low.   You should substitute TWO of your usual THREE meals daily with TWO Formula 1 Smoothies until your BMI is normal (between 20.7 and 26.4).   Then go on maintenance."
Obesity = "N"

Case 27.81 To 31.1
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are OVERWEIGHT.  Moderate risk of Obesity Co-morbidities exists.  You should substitute TWO of your usual THREE meals daily with TWO Formula 1 Smoothies until your BMI is normal (between 20.7 and 26.4).   Then go on maintenance."
Obesity = "Y"
Case 31.11 To 45.4
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are SEVERELY OVERWEIGHT.  High risk of Obesity Co-morbidities exists.  You should substitute TWO of your usual THREE meals daily with TWO Formula 1 Smoothies until your BMI is normal (between 20.7 and 26.4).   Then go on maintenance.   You need to be under a Physician's Care as soon as possible."
Obesity = "Y"
Case Is > 45.4
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are  MORBIDLY OBESE.  VERY HIGH RISK for Obesity Co-morbidities exists.  You should substitute TWO of your usual THREE meals daily with TWO Formula 1 Smoothies until your BMI is normal (between 20.7 and 26.4).   Then go on maintenance.   You need to be under a Physician's Care as soon as possible."
Obesity = "Y"
End Select

Else
Select Case BmiValue

Case Is < 19.1
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are UNDERWEIGHT. The lower the BMI the greater the risk of Malnutrition.   You should supplement each of your THREE meals daily with a Formula 1 Smoothie until your BMI is normal (between 20.7 and 26.4).   Then go on maintenance."
Obesity = "N"

Case 19.11 To 25.8
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are at a NORMAL WEIGHT.  Your risk of Co-morbidities is very low.   You should substitute ONE of your usual THREE meals daily with a ONE Formula 1 Smoothie and maintain your BMI as is (between 20.7 and 26.4).   Congratulations!   Stay on maintenance."
Obesity = "N"

Case 25.81 To 27.3
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are  MARGINALLY OVERWEIGHT.  Some risk of Obesity Co-morbidities exists but they are low.   You should substitute TWO of your usual THREE meals daily with TWO Formula 1 Smoothies until your BMI is normal (between 20.7 and 26.4).   Then go on maintenance."
Obesity = "N"

Case 27.31 To 32.2
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are  OVERWEIGHT.  Moderate risk of Obesity Co-morbidities exists.  You should substitute TWO of your usual THREE meals daily with TWO Formula 1 Smoothies until your BMI is normal (between 20.7 and 26.4).   Then go on maintenance."
Obesity = "Y"
Case 32.21 To 44.8
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are  SEVERELY OVERWEIGHT.  High risk of Obesity Co-morbidities exists.  You should substitute TWO of your usual THREE meals daily with TWO Formula 1 Smoothies until your BMI is normal (between 20.7 and 26.4).   Then go on maintenance.   You need to be under a Physician's Care as soon as possible."
Obesity = "Y"
Case Is > 44.8
BMIcaption = vbCr & vbLf & vbCr & vbLf & "You are  MORBIDLY OBESE.  VERY HIGH RISK for Obesity Co-morbidities exists.  You should substitute TWO of your usual THREE meals daily with TWO Formula 1 Smoothies until your BMI is normal (between 20.7 and 26.4).   Then go on maintenance.   You need to be under a Physician's Care as soon as possible."
Obesity = "Y"
End Select
End If

End Sub

Private Sub CAD_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub CAD_LostFocus()
If UCase(LTrim(RTrim(CAD))) <> "Y" And UCase(LTrim(RTrim(CAD))) <> "N" Then
    CAD = ""
Else
CAD = UCase((LTrim(RTrim(CAD))))
End If
End Sub

Private Sub Check1_Click(Index As Integer)
Select Case Index
Case 0
    Check1(1).Value = 0
    Check1(2).Value = 0
Case 1
    Check1(0).Value = 0
    Check1(2).Value = 0
Case 2
    Check1(1).Value = 0
    Check1(0).Value = 0
End Select
End Sub


Private Sub Check2_Click(Index As Integer)
Select Case Index
Case 0
    Check2(1).Value = 0
    Check2(2).Value = 0
Case 1
    Check2(0).Value = 0
    Check2(2).Value = 0
Case 2
    Check2(1).Value = 0
    Check2(0).Value = 0
End Select

End Sub

Private Sub Combo1_Click()

If Val(Inny) = 0 Then
    Exit Sub
    Inny.SetFocus
End If

Select Case Combo1.Text

Case "Kilograms to Pounds"
Outy = Str(Val(Inny) * 2.2)

Case "Pounds to Kilograms"
Outy = Str(Val(Inny) / 2.2)

Case "Centimeters to Inches"
Outy = Str(Val(Inny) / 2.54)

Case "Inches to Centimeters"
Outy = Str(Val(Inny) * 2.54)

Case "mmol/L to mg/dl"
Outy = Str(Val(Inny) / 0.0259)

Case "mg/dl to mmol/L"
Outy = Str(Val(Inny) * 0.0259)

Case "gm/L to mg/dl"
Outy = Str(Val(Inny) / 100)

Case "mg/dl to gm/L"
Outy = Str(Val(Inny) * 100)

End Select

End Sub

Private Sub Command1_Click()


End Sub

Private Sub Command2_Click()
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Dim Mortise As Integer
Dim MortiseRisk As String
Dim E10, E11, E12 As Integer
Dim Fib, Cholesterol1 As Double
Dim FibRisk As String
Dim HDLrisk As String
Dim HDLriskV As Double
Dim LDLrisk As String
Dim LDLriskV As Double
Dim NCEP As Integer
Dim NCEPs As String
Dim NCEP1 As Integer
Dim NCEPs1 As String
Dim ATP As Integer
Dim ATPs As String
Dim ATPx As Integer
ATPx = 0
ATP = 0
ATPs = ""
NCEP = 0
NCEPs = ""
NCEP1 = 0
NCEPs1 = ""
LDLrisk = ""
LDLriskV = 0
HDLrisk = ""
Fibrinogen = 0
FibRisk = ""
AHA = 0
AHArisk = ""
Mortise = 0
MortiseRisk = ""
E10 = 0

If Text1(1) = "" Or Text1(2) = "" Or Text1(13) = "" Or Text2 = "" Or Text1(6) = "" Or Text1(7) = "" Or Text1(8) = "" _
Or Text1(9) = "" Or Text1(10) = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text8 = "" Or Text10 = "" _
Or Text11 = "" Or CAD = "" Or Hypertension = "" Or Hyperlipidemia = "" Or FhxCAD = "" Or Obesity = "" Then
    Msg = "You MUST at the very least Complete Gender, Age, Weight, Height and ANSWER ALL Y/N Questions."
    Style = vbCritical
    Title = "Input Error"   ' Define title.
    Help = "DEMO.HLP"   ' Define Help file.
    Ctxt = 1000   ' Define topic
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
Exit Sub
End If

If Text1(0) = "M" Then   'gender
    E10 = 1
    Text3.Text = "N"
    Text4.Text = "N"
    Select Case Val(Text1(1)) 'Age
    Case Is >= 45
        ATP = 1
    Case 35 To 44.9
        NCEP1 = 1
    Case 45 To 54.9
        NCEP1 = 2
    Case 55 To 64.9
        NCEP1 = 3
    Case 65 To 74.9
        NCEP1 = 4
    Case 75 To 84.9
        NCEP1 = 5
    Case 85 To 94.9
        NCEP1 = 6
    Case 95 To 104.9
        NCEP1 = 7
    Case Is >= 105
        NCEP1 = 8
    Case Is < 40
        Mortise = 3
    Case Is >= 45
        NCEP = 1
    Case 40 To 55
        Mortise = 6
    Case Is > 55
        Mortise = 9
    End Select
    
    Select Case Val(Text1(1)) 'Age
    Case Is < 35
        AHA = 0
    Case 35.1 To 40
        AHA = 1
    Case 40.1 To 49
        AHA = 2
    Case 49.1 To 54
        AHA = 3
    Case Is > 54
        AHA = 4
    End Select
    
Else  'Female

    If Text3 = "Y" And Text4 = "Y" Then NCEP = 1
    Select Case Val(Text1(1)) 'Age
    Case Is >= 55
        ATP = 1
    Case 35 To 44.9
        NCEP1 = 0
    Case 45 To 54.9
        NCEP1 = 1
    Case 55 To 64.9
        NCEP1 = 2
    Case 65 To 74.9
        NCEP1 = 3
    Case 75 To 84.9
        NCEP1 = 4
    Case 85 To 94.9
        NCEP1 = 5
    Case 95 To 104.9
        NCEP1 = 6
    Case Is >= 105
        NCEP1 = 7

    Case Is <= 50
        Mortise = 3
    Case Is >= 55
        NCEP = 1
    Case 50 To 65
        Mortise = 6
    Case Is > 65
        Mortise = 9
    End Select
    
    Select Case Val(Text1(1))
    Case Is < 42
        AHA = 0
    Case 42.1 To 45
        AHA = 1
    Case 45.1 To 55
        AHA = 3
    Case 45.1 To 74
        AHA = 3
    Case Is > 74
        AHA = 4
    End Select
End If
If Text1(11) = "" Or Text1(12) = "" Then
    Msg = "The very minimal interpretation requires a Total Cholesterol and an HDL-Cholesterol."
    Style = vbCritical
    Title = "Input Error"   ' Define title.
    Help = "DEMO.HLP"   ' Define Help file.
    Ctxt = 1000   ' Define topic
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
Exit Sub
End If


If Text6 = "Y" Or Text7 = "Y" Then NCEP = NCEP + 1


If Text1(6) = "Y" Then AHA = AHA + 2 'FHx of Heart ds or MI prior to age of 60? (Y/N)

If Text1(7) = "Y" Then AHA = AHA + 1 'Rarely exercise or do anything physical? (Y/N)

If Text1(8) = "Y" Then 'Do you smoke? (Y/N)
    AHA = AHA + 1
    Mortise = Mortise + 1
    E11 = 1
    NCEP = NCEP + 1
    NCEP1 = NCEP1 + 1
    ATP = ATP + 1
End If

If Text1(9) = "Y" And Text1(0) = "M" Then
    AHA = AHA + 1 'Diabetic? (Y/N)
    Mortise = Mortise + 2
    NCEP = NCEP + 1
    NCEP1 = NCEP1 + 1
End If
If Text1(9) = "Y" And Text1(0) = "F" Then
    AHA = AHA + 2 'female diabetic
    Mortise = Mortise + 2
    NCEP = NCEP + 1
    NCEP1 = NCEP1 + 1
End If

If Text10 = "Y" Or Text11 = "Y" Then ATP = ATP + 1

If Text5 = "Y" Or Text1(10) = "Y" Then
    NCEP = NCEP + 1
    NCEP1 = NCEP1 + 1
End If
If Val(Syst) >= 140 And Val(diastolic) >= 9 Then
    ATP = ATP + 1
Else
    If Text1(10) = "Y" Then ATP = ATP + 1
End If

If Val(Text1(12)) >= 60 Then
    NCEP = NCEP - 1
    NCEP1 = NCEP1 - 1
    ATP = ATP - 1
End If
If Val(Text1(12)) < 40 And Text1(12) <> "" Then
    ATP = ATP + 1
End If

If Val(Text1(12)) < 35 And Text1(12) <> "" Then
    NCEP = NCEP + 1
    NCEP1 = NCEP1 + 1
End If

If Val(LDL) >= 200 Then NCEP1 = NCEP1 + 1

If Hypertension = "Y" Then
    Mortise = Mortise + 1
    E12 = 1
End If
If Hyperlipidemia = "Y" Then
    Mortise = Mortise + 1
End If
If FhxCAD = "Y" Then
    Mortise = Mortise + 1
End If
If Obesity = "Y" Then
    Mortise = Mortise + 1
End If
If Check1(0).Value = 1 Then Mortise = Mortise + 5
If Check1(1).Value = 1 Then Mortise = Mortise + 3
If Check1(2).Value = 1 Then Mortise = Mortise + 1
If Check2(0).Value = 1 Then Mortise = Mortise + 0
If Check2(1).Value = 1 Then Mortise = Mortise - 3
If Check2(2).Value = 1 Then Mortise = Mortise + 3




If Text1(10) = "N" Then 'Current BP meds? (Y/N)
    If Val(Syst1) > 170 Then
        AHA = AHA + 2
    End If
    If Val(Syst1) >= 140 And Val(Syst1) <= 170 Then
    AHA = AHA + 1
    End If
Else
    AHA = AHA + 1
End If
Select Case Val(Text1(11)) 'Current Total Cholesterol level, in mg/dL
    Case 240 To 315
        AHA = AHA + 1
    Case Is > 315
        AHA = AHA + 2
End Select
'=IF(B23>=60,-1,IF(B23>38,0,IF(B23>=30,1,2)))
Select Case Val(Text1(12)) 'Current HDL cholesterol level, in mg/dL
    Case Is < 30
        AHA = AHA + 2
    Case 30 To 38
        AHA = AHA + 1
    Case Is >= 60
        AHA = AHA - 1
End Select

If AHA >= 4 Then
AHArisk = "ABOVE AVERAGE risk of a first heart attack relative to the general adult population"
Else
AHArisk = "AVERAGE risk of a first heart attack relative to the general adult population"
End If
If Mortise <= 8 Then MortiseRisk = "Low"
If Mortise > 8 And Mortise <= 15 Then MortiseRisk = "Intermediate"
If Mortise > 15 Then MortiseRisk = "High"
Fib = (0.05 * Val(Text1(1))) + (0.1 * Val(Text1(11)) / Val(Text1(12))) + (0.5 * E10) + (0.002857 * Val(Fibrinogen)) + (0.5 * E11) + (0.7 * E12)
FibRisk = Str(Int(Fib * 100) / 100)


Interpret = "The Total number of AHA points is: " & Str(AHA) & " out of 21." & vbCr & vbLf & vbCr & vbLf & AHArisk & " (American Heart Association)."

If Text2 = "Y" Then
    Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "The Clinical Score of Mortise for Coronary Artery Disease is : " & Mortise & vbCr & vbLf & vbCr & vbLf & "The Mortise Risk for the presence of CAD is " & MortiseRisk
End If

If Val(Fibrinogen) <> 0 Then
    Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "The Fibrinogen (Schermund) based Overall risk for Coronary Atherosclerotic Disease is : " & FibRisk
End If

If Triglyceride <> "" Then
    If Val(Triglyceride) <= 400 Then
        Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "The Calculated VLDL is : " & Val(triglycerides) / 5
    Else
        Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "One can not accurately Calculate the VLDL as the Triglycerides exceed 400 mg/dl.   Measure it directly."
    End If
End If
If Triglyceride <> "" Then
    If Val(Triglyceride) <= 400 Then
        Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "The Calculated LDL is : " & Val(Text1(11)) - (Val(triglycerides) / 5) - Val(Text1(12)) & " (Cholesterol - VLDL - HDL)"
    Else
        Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "One can not accurately Calculate the LDL as the Triglycerides exceed 400 mg/dl.   Measure LDL directly."
    End If
End If

If LDL <> "" Then
    Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "The measured LDL Cholesterol is " & LDL & " mg/dl"
End If

If Text1(0) = "M" Then
Select Case Val(Text1(12))
    Case Is > 65
        HDLrisk = "Very Low Risk"
    Case 25 To 65 '
        HDLriskV = (0.000658 * (Val(Text1(12)) ^ 2)) - (0.0979 * Val(Text1(12))) + 4.085
        HDLrisk = Str(HDLriskV) & " times the Average Risk"
    Case Is < 25
        HDLrisk = "Very High Risk"
End Select
Else
Select Case Val(Text1(12))
    Case Is > 70
        HDLrisk = "Very Low Risk"
    Case 40 To 70
        HDLriskV = (0.0008095 * (Val(Text1(12)) ^ 2)) - (0.136905 * Val(Text1(12))) + 6.1
        HDLrisk = Str(HDLriskV) & " times the Average Risk"
    Case Is < 40
        HDLrisk = "Very High Risk"
End Select
End If

If Text1(12) <> "" And Text1(11) <> "" Then

Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "Based upon the HDL, the following situation exists: " & HDLrisk
HDLrisk = ""

If Text1(0) = "M" Then
Select Case Val(Text1(11)) / Val(Text1(12))
    Case Is > 23.3
        HDLrisk = "Three Times Average Risk"
    Case 9.51 To 23.3
        HDLrisk = "Two Times Average Risk"
    Case 4.91 To 9.5
        HDLrisk = "Average Risk"
    Case Is <= 4.9
        HDLrisk = "50% Average Risk"
End Select
Else
Select Case Val(Text1(11)) / Val(Text1(12))
    Case Is > 10.9
        HDLrisk = "Three Times Average Risk"
    Case 6.91 To 10.9
        HDLrisk = "Two Times Average Risk"
    Case 4.31 To 6.9
        HDLrisk = "Average Risk"
    Case Is <= 4.3
        HDLrisk = "50% Average Risk"
End Select
End If
Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "The Total Cholesterol to HDL Ratio is: " & Str((Int((Val(Text1(11)) / Val(Text1(12))) * 100)) / 100) & " which indicates a " & HDLrisk
End If

If Val(LDL) <> 0 And Val(trigLygerides) <> 0 And Val(Text1(12)) <> 0 And Val(Text1(11)) <> 0 Then

If LDL <> "" Then
LDLriskV = (Int((Val(LDL) / Val(Text1(12))) * 100)) / 100
Else
LDLriskV = (Int(((Val(Text1(11)) - (Val(triglycerides) / 5) - Val(Text1(12))) / Val(Text1(12))) * 100)) / 100
End If
Select Case LDLriskV
Case Is < 3
    LDLrisk = "Low Risk"
Case 3 To 6
    LDLrisk = "Average Risk"
Case Is > 6
    LDLrisk = "High Risk"
End Select
Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "The LDL Cholesterol to HDL Cholesterol Ratio is " & LDLriskV & " this represents " & LDLrisk
End If
'M =IF(B7="F","NA",IF(AND(B8>=20,B8<=34,B22<2,B17>=220),"Yes",IF(AND(B8>=20,B8<=34,B22>=2,B17>=190),"Yes",
'IF(AND(B8>=35,B22<2,B17>=190),"Yes",IF(AND(B8>=35,B22>=2,B17>=160),"Yes","No")))))

'F=IF(B7="M","NA",IF(AND(B9="Y",B22<2,B17>=190),"Yes",IF(AND(B9="Y",B22>=2,B17>=160),"Yes",
'IF(AND(B8>=20,B9="N",B22<2,B17>=220),"Yes",IF(AND(B8>=20,B9="N",B22>=2,B17>=190),"Yes","No")))))
If Text1(0) = "M" Then
If Val(Text1(1)) >= 20 And Val(Text1(1)) <= 34 And NCEP < 2 And Val(LDL) >= 220 Then NCEPs = "Treatment is indicated"
If Val(Text1(1)) >= 20 And Val(Text1(1)) <= 34 And NCEP >= 2 And Val(LDL) >= 190 Then NCEPs = "Treatment is indicated"
If Val(Text1(1)) >= 35 And NCEP < 2 And Val(LDL) >= 190 Then NCEPs = "Treatment is indicated"
If Val(Text1(1)) >= 35 And NCEP >= 2 And Val(LDL) >= 160 Then NCEPs = "Treatment is indicated"
Else
If Text3 = "Y" And NCEP < 2 And Val(LDL) >= 190 Then NCEPs = "Treatment is indicated"
If Text3 = "Y" And NCEP >= 2 And Val(LDL) >= 160 Then NCEPs = "Treatment is indicated"
If Val(Text1(1)) >= 20 And Text3 = "N" And NCEP < 2 And Val(LDL) >= 220 Then NCEPs = "Treatment is indicated"
If Val(Text1(1)) >= 20 And Text3 = "N" And NCEP >= 2 And Val(LDL) >= 190 Then NCEPs = "Treatment is indicated"
End If
If NCEPs = "" Then NCEPs = "Treatment is not indicated"
Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "You have " & NCEP & " cardiac risk factors based upon the National Cholesterol Educational Program II (NCEPII) criteria.  " & NCEPs & " based upon your values."
If Val(LDL) >= 160 And NCEP1 >= 3 Then
    NCEPs1 = "Treatment is indicated"
Else
    NCEPs1 = "Treatment is not indicated"
End If
Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "You have " & NCEP1 & " cardiac risk factors based upon the National Cholesterol Educational Program II (NCEPII) Revised criteria.  " & NCEPs & " based upon your values."

If Val(Text1(11)) >= 240 Then
    ATPs = "The total cholesterol is HIGH.   "
Else
    If Val(Text1(11)) >= 200 Then
            ATPs = "The total cholesterol is BORDERLINE HIGH.   "
    Else
            ATPs = "The total cholesterol is DESIRABLE.   "
    End If
End If
If Val(Text1(12)) >= 60 Then
    ATPs = ATPs & "The HDL-Cholesterol is HIGH and (GOOD).   "
Else
    If Val(Text1(12)) >= 40 Then
        ATPs = ATPs & "The HDL-Cholesterol is INTERMEDIATE.   "
    Else
            ATPs = ATPs & "The HDL-Cholesterol is LOW and (NOT GOOD).   "
    End If
End If
'========================================================================
If Val(LDL) = 0 Then
    Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & ATPs
Else
'IF(B11>=190,"very high",IF(B11>=160,"high",IF(B11>=130,"borderline high",IF(B11>=100,"above optimal","optimal")))))
If Val(LDL) >= 160 Then
    ATPs = ATPs & "The LDL-Cholesterol is HIGH.   "
Else
    If Val(LDL) >= 130 Then
        ATPs = ATPs & "The LDL-Cholesterol is BORDERLINE HIGH.   "
    Else
        If Val(LDL) >= 100 Then
            ATPs = ATPs & "The LDL-Cholesterol is ABOVE OPTIMAL.   "
        Else
                ATPs = ATPs & "The LDL-Cholesterol is OPTIMAL.   "
        End If
    End If
End If


ATPs = ATPs & "The Number of Coronary Risk Factors from the Adult Treatment Panel III criteria is/are: " & ATP & ".   "
'(B17="Y",B18="Y"),"< 100",IF(B28>=2,"< 130","< 160")))
If CAD = "Y" Or Text1(9) = "Y" Then
    ATPs = ATPs & "The LDL Cholesterol goal in mg/dl is less than 100.   "
    ATPx = 99
Else
    If ATP >= 2 Then
        ATPs = ATPs & "The LDL Cholesterol goal in mg/dl is less than 130.   "
        ATPx = 129
    Else
            ATPs = ATPs & "The LDL Cholesterol goal in mg/dl is less than 160.   "
            ATPx = 159
    End If
End If
If Val(LDL) < ATPx Then
        ATPs = ATPs & "The LDL Cholesterol does not require reduction.   "
Else
        ATPs = ATPs & "The LDL Cholesterol needs to be reduced by at least " & Val(LDL) - ATPx & "."
End If
If Syst1 <> "" Or diastolic <> "" Then
Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & ATPs
End If
End If
'=======================================================================================================
Interpret = Interpret & vbCr & vbLf & vbCr & vbLf & "Your BMI is " & BmiValue & ". " & BMIcaption
End Sub

Private Sub FhxCAD_LostFocus()
If UCase(LTrim(RTrim(FhxCAD))) <> "Y" And UCase(LTrim(RTrim(FhxCAD))) <> "N" Then
    FhxCAD = ""
Else
FhxCAD = UCase((LTrim(RTrim(FhxCAD))))
End If
End Sub

Private Sub FhxCAD_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub Fibrinogen_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "#" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeySpace Then
KeyAscii = 0
End If

End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Combo1.AddItem "Kilograms to Pounds"
Combo1.AddItem "Pounds to Kilograms"
Combo1.AddItem "Centimeters to Inches"
Combo1.AddItem "Inches to Centimeters"
Combo1.AddItem "mmol/L to mg/dl"
Combo1.AddItem "mg/dl to mmol/L"
Combo1.AddItem "gm/L to mg/dl"
Combo1.AddItem "mg/dl to gm/L"
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "This Module is Intended as an Educational Device.   It is not warranteed for accuracy and is not designed to substitute for a clinical evaluation by your personal physician.   By clicking YES below, you agree with these terms."
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "Disclaimer"   ' Define title.
    Help = "DEMO.HLP"   ' Define Help file.
    Ctxt = 1000   ' Define topic
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbNo Then   ' User chose Yes.
   Unload Me   ' Perform some action.
   Exit Sub
End If
 
End Sub

Private Sub Hyperlipidemia_LostFocus()
If UCase(LTrim(RTrim(Hyperlipidemia))) <> "Y" And UCase(LTrim(RTrim(Hyperlipidemia))) <> "N" Then
    Hyperlipidemia = ""
Else
Hyperlipidemia = UCase((LTrim(RTrim(Hyperlipidemia))))
End If

End Sub

Private Sub Hyperlipidemia_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub Hypertension_LostFocus()
If UCase(LTrim(RTrim(Hypertension))) <> "Y" And UCase(LTrim(RTrim(Hypertension))) <> "N" Then
    Hypertension = ""
Else
Hypertension = UCase((LTrim(RTrim(Hypertension))))
End If
End Sub

Private Sub Hypertension_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub Obesity_LostFocus()
If UCase(LTrim(RTrim(Obesity))) <> "Y" And UCase(LTrim(RTrim(Obesity))) <> "N" Then
    Obesity = ""
Else
FhxCAD = UCase((LTrim(RTrim(Obesity))))
End If
End Sub

Private Sub Obesity_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

Case 1, 2
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "#" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeySpace Then
KeyAscii = 0
End If

End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If UCase(LTrim(RTrim(Text1(0)))) <> "F" And UCase(LTrim(RTrim(Text1(0)))) <> "M" Then
    Text1(0) = ""
    Text1(0).SetFocus
Else
Text1(0) = UCase((LTrim(RTrim(Text1(0)))))
If Text1(0) = "M" Then
    Text3.Text = "N"
    Text4.Text = "N"
End If
End If

If Text1(2) <> "" And Text1(3) <> "" Then
    Bmi
    Text1(19) = Str(BmiValue)
End If

If UCase(LTrim(RTrim(Text1(6)))) <> "Y" And UCase(LTrim(RTrim(Text1(6)))) <> "N" Then
    Text1(6) = ""
Else
Text1(6) = UCase((LTrim(RTrim(Text1(6)))))
End If

If UCase(LTrim(RTrim(Text1(7)))) <> "Y" And UCase(LTrim(RTrim(Text1(7)))) <> "N" Then
    Text1(7) = ""
Else
Text1(7) = UCase((LTrim(RTrim(Text1(7)))))
End If

If UCase(LTrim(RTrim(Text1(8)))) <> "Y" And UCase(LTrim(RTrim(Text1(8)))) <> "N" Then
    Text1(8) = ""
Else
Text1(8) = UCase((LTrim(RTrim(Text1(8)))))
End If

If UCase(LTrim(RTrim(Text1(9)))) <> "Y" And UCase(LTrim(RTrim(Text1(9)))) <> "N" Then
    Text1(9) = ""
Else
Text1(9) = UCase((LTrim(RTrim(Text1(9)))))
End If

If UCase(LTrim(RTrim(Text1(10)))) <> "Y" And UCase(LTrim(RTrim(Text1(10)))) <> "N" Then
    Text1(10) = ""
Else
Text1(10) = UCase((LTrim(RTrim(Text1(10)))))
Hypertension.Text = "Y"
End If

If UCase(LTrim(RTrim(Text1(13)))) <> "Y" And UCase(LTrim(RTrim(Text1(13)))) <> "N" Then
    Text1(13) = ""
Else
Text1(13) = UCase((LTrim(RTrim(Text1(13)))))
End If

'If UCase(LTrim(RTrim(Text1(14)))) <> "Y" And UCase(LTrim(RTrim(Text1(14)))) <> "N" Then
'    Text1(15) = ""
'Else
'Text1(15) = UCase((LTrim(RTrim(Text1(14)))))
'End If

'If UCase(LTrim(RTrim(Text1(15)))) <> "Y" And UCase(LTrim(RTrim(Text1(15)))) <> "N" Then
'    Text1(15) = ""
'Else
'Text1(15) = UCase((LTrim(RTrim(Text1(15)))))
'End If

'If UCase(LTrim(RTrim(Text1(16)))) <> "Y" And UCase(LTrim(RTrim(Text1(16)))) <> "N" Then
'    Text1(16) = ""
'Else
'Text1(16) = UCase((LTrim(RTrim(Text1(16)))))
'End If


'If UCase(LTrim(RTrim(Text1(17)))) <> "Y" And UCase(LTrim(RTrim(Text1(17)))) <> "N" Then
'    Text1(17) = ""
'Else
'Text1(17) = UCase((LTrim(RTrim(Text1(17)))))
'End If

End Sub

Private Sub Text10_LostFocus()
If UCase(LTrim(RTrim(Text10))) <> "Y" And UCase(LTrim(RTrim(Text10))) <> "N" Then
    Text10 = ""
Else
Text10 = UCase((LTrim(RTrim(Text10))))
Text8 = "Y"
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub Text11_LostFocus()
If UCase(LTrim(RTrim(Text11))) <> "Y" And UCase(LTrim(RTrim(Text11))) <> "N" Then
    Text11 = ""
Else
Text11 = UCase((LTrim(RTrim(Text11))))
Text6 = "Y"
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub Text2_LostFocus()
If UCase(LTrim(RTrim(Text2))) <> "Y" And UCase(LTrim(RTrim(Text2))) <> "N" Then
    Text2 = ""
Else
Text2 = UCase((LTrim(RTrim(Text2))))
End If
If Text2 = "Y" Then
    Frame1.Enabled = True
Else
    Frame1.Enabled = False
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub Text3_LostFocus()
If UCase(LTrim(RTrim(Text3))) <> "Y" And UCase(LTrim(RTrim(Text3))) <> "N" Then
    Text3 = ""
Else
Text3 = UCase((LTrim(RTrim(Text3))))
If Text3 = "N" Then Text4 = "N"
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub Text4_LostFocus()
If UCase(LTrim(RTrim(Text4))) <> "Y" And UCase(LTrim(RTrim(Text4))) <> "N" Then
    Text4 = ""
Else
Text4 = UCase((LTrim(RTrim(Text4))))
End If
End Sub

Private Sub Text5_LostFocus()
If UCase(LTrim(RTrim(Text5))) <> "Y" And UCase(LTrim(RTrim(Text5))) <> "N" Then
    Text5 = ""
Else
Text5 = UCase((LTrim(RTrim(Text5))))
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub Text6_LostFocus()
If UCase(LTrim(RTrim(Text6))) <> "Y" And UCase(LTrim(RTrim(Text6))) <> "N" Then
    Text6 = ""
Else
Text6 = UCase((LTrim(RTrim(Text6))))
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

Private Sub Text8_LostFocus()
If UCase(LTrim(RTrim(Text8))) <> "Y" And UCase(LTrim(RTrim(Text8))) <> "N" Then
    Text8 = ""
Else
Text8 = UCase((LTrim(RTrim(Text8))))
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) Like "[a-zA-Z]" <> True And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyDelete Then
KeyAscii = 0
End If

End Sub

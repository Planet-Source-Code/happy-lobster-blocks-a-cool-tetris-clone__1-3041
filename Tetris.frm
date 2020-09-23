VERSION 5.00
Begin VB.Form Tetris 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blocks!"
   ClientHeight    =   6345
   ClientLeft      =   690
   ClientTop       =   375
   ClientWidth     =   8175
   FillColor       =   &H0000C0C0&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tetris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6345
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   720
      Top             =   2400
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Next:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Blocks!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by Anton Gustavsson"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Styled by Happy Crab"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   0
      Left            =   360
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   1
      Left            =   600
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   2
      Left            =   840
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   3
      Left            =   1080
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   4
      Left            =   360
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   5
      Left            =   600
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   6
      Left            =   840
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   7
      Left            =   1080
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   8
      Left            =   360
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   9
      Left            =   600
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   10
      Left            =   840
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Index           =   11
      Left            =   1080
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF80FF&
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   735
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   4560
      TabIndex        =   0
      Top             =   2280
      Width           =   300
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   1665
      X2              =   4105
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   4110
      X2              =   4110
      Y1              =   120
      Y2              =   6140
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   249
      Left            =   3840
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   248
      Left            =   3600
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   247
      Left            =   3360
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   246
      Left            =   3120
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   245
      Left            =   2880
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   244
      Left            =   2640
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   243
      Left            =   2400
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   242
      Left            =   2160
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   241
      Left            =   1920
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   240
      Left            =   1680
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   239
      Left            =   3840
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   238
      Left            =   3600
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   237
      Left            =   3360
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   236
      Left            =   3120
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   235
      Left            =   2880
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   234
      Left            =   2640
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   233
      Left            =   2400
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   232
      Left            =   2160
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   231
      Left            =   1920
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   230
      Left            =   1680
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   229
      Left            =   3840
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   228
      Left            =   3600
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   227
      Left            =   3360
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   226
      Left            =   3120
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   225
      Left            =   2880
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   224
      Left            =   2640
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   223
      Left            =   2400
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   222
      Left            =   2160
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   221
      Left            =   1920
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   220
      Left            =   1680
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   219
      Left            =   3840
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   218
      Left            =   3600
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   217
      Left            =   3360
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   216
      Left            =   3120
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   215
      Left            =   2880
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   214
      Left            =   2640
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   213
      Left            =   2400
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   212
      Left            =   2160
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   211
      Left            =   1920
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   210
      Left            =   1680
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   209
      Left            =   3840
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   208
      Left            =   3600
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   207
      Left            =   3360
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   206
      Left            =   3120
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   205
      Left            =   2880
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   204
      Left            =   2640
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   203
      Left            =   2400
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   202
      Left            =   2160
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   201
      Left            =   1920
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   200
      Left            =   1680
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   199
      Left            =   3840
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   198
      Left            =   3600
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   197
      Left            =   3360
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   196
      Left            =   3120
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   195
      Left            =   2880
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   194
      Left            =   2640
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   193
      Left            =   2400
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   192
      Left            =   2160
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   191
      Left            =   1920
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   190
      Left            =   1680
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   189
      Left            =   3840
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   188
      Left            =   3600
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   187
      Left            =   3360
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   186
      Left            =   3120
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   185
      Left            =   2880
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   184
      Left            =   2640
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   183
      Left            =   2400
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   182
      Left            =   2160
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   181
      Left            =   1920
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   180
      Left            =   1680
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   179
      Left            =   3840
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   178
      Left            =   3600
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   177
      Left            =   3360
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   176
      Left            =   3120
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   175
      Left            =   2880
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   174
      Left            =   2640
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   173
      Left            =   2400
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   172
      Left            =   2160
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   171
      Left            =   1920
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   170
      Left            =   1680
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   169
      Left            =   3840
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   168
      Left            =   3600
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   167
      Left            =   3360
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   166
      Left            =   3120
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   165
      Left            =   2880
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   164
      Left            =   2640
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   163
      Left            =   2400
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   162
      Left            =   2160
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   161
      Left            =   1920
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   160
      Left            =   1680
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   159
      Left            =   3840
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   158
      Left            =   3600
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   157
      Left            =   3360
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   156
      Left            =   3120
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   155
      Left            =   2880
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   154
      Left            =   2640
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   153
      Left            =   2400
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   152
      Left            =   2160
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   151
      Left            =   1920
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   150
      Left            =   1680
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   149
      Left            =   3840
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   148
      Left            =   3600
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   147
      Left            =   3360
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   146
      Left            =   3120
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   145
      Left            =   2880
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   144
      Left            =   2640
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   143
      Left            =   2400
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   142
      Left            =   2160
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   141
      Left            =   1920
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   140
      Left            =   1680
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   139
      Left            =   3840
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   138
      Left            =   3600
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   137
      Left            =   3360
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   136
      Left            =   3120
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   135
      Left            =   2880
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   134
      Left            =   2640
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   133
      Left            =   2400
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   132
      Left            =   2160
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   131
      Left            =   1920
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   130
      Left            =   1680
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   129
      Left            =   3840
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   128
      Left            =   3600
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   127
      Left            =   3360
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   126
      Left            =   3120
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   125
      Left            =   2880
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   124
      Left            =   2640
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   123
      Left            =   2400
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   122
      Left            =   2160
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   121
      Left            =   1920
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   120
      Left            =   1680
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   119
      Left            =   3840
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   118
      Left            =   3600
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   117
      Left            =   3360
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   116
      Left            =   3120
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   115
      Left            =   2880
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   114
      Left            =   2640
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   113
      Left            =   2400
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   112
      Left            =   2160
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   111
      Left            =   1920
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   110
      Left            =   1680
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   109
      Left            =   3840
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   108
      Left            =   3600
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   107
      Left            =   3360
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   106
      Left            =   3120
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   105
      Left            =   2880
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   104
      Left            =   2640
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   103
      Left            =   2400
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   102
      Left            =   2160
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   101
      Left            =   1920
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   100
      Left            =   1680
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   99
      Left            =   3840
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   98
      Left            =   3600
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   97
      Left            =   3360
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   96
      Left            =   3120
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   95
      Left            =   2880
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   94
      Left            =   2640
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   93
      Left            =   2400
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   92
      Left            =   2160
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   91
      Left            =   1920
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   90
      Left            =   1680
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   89
      Left            =   3840
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   88
      Left            =   3600
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   87
      Left            =   3360
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   86
      Left            =   3120
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   85
      Left            =   2880
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   84
      Left            =   2640
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   83
      Left            =   2400
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   82
      Left            =   2160
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   81
      Left            =   1920
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   80
      Left            =   1680
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   79
      Left            =   3840
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   78
      Left            =   3600
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   77
      Left            =   3360
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   76
      Left            =   3120
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   75
      Left            =   2880
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   74
      Left            =   2640
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   73
      Left            =   2400
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   72
      Left            =   2160
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   71
      Left            =   1920
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   70
      Left            =   1680
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   69
      Left            =   3840
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   68
      Left            =   3600
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   67
      Left            =   3360
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   66
      Left            =   3120
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   65
      Left            =   2880
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   64
      Left            =   2640
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   63
      Left            =   2400
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   62
      Left            =   2160
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   61
      Left            =   1920
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   60
      Left            =   1680
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   59
      Left            =   3840
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   58
      Left            =   3600
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   57
      Left            =   3360
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   56
      Left            =   3120
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   55
      Left            =   2880
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   54
      Left            =   2640
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   53
      Left            =   2400
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   52
      Left            =   2160
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   51
      Left            =   1920
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   50
      Left            =   1680
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   49
      Left            =   3840
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   48
      Left            =   3600
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   47
      Left            =   3360
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   46
      Left            =   3120
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   45
      Left            =   2880
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   44
      Left            =   2640
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   43
      Left            =   2400
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   42
      Left            =   2160
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   41
      Left            =   1920
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   40
      Left            =   1680
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   39
      Left            =   3840
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   38
      Left            =   3600
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   37
      Left            =   3360
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   36
      Left            =   3120
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   35
      Left            =   2880
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   34
      Left            =   2640
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   33
      Left            =   2400
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   32
      Left            =   2160
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   31
      Left            =   1920
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   30
      Left            =   1680
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   29
      Left            =   3840
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   28
      Left            =   3600
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   27
      Left            =   3360
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   26
      Left            =   3120
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   25
      Left            =   2880
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   24
      Left            =   2640
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   23
      Left            =   2400
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   22
      Left            =   2160
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   21
      Left            =   1920
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   20
      Left            =   1680
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   19
      Left            =   3840
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   18
      Left            =   3600
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   17
      Left            =   3360
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   16
      Left            =   3120
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   15
      Left            =   2880
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   14
      Left            =   2640
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   13
      Left            =   2400
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   12
      Left            =   2160
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   11
      Left            =   1920
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   10
      Left            =   1680
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   9
      Left            =   3840
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   8
      Left            =   3600
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   7
      Left            =   3360
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   6
      Left            =   3120
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   5
      Left            =   2880
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   4
      Left            =   2640
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   3
      Left            =   2400
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   2
      Left            =   2160
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   1
      Left            =   1920
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   0
      Left            =   1680
      Top             =   5880
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   1665
      X2              =   1665
      Y1              =   120
      Y2              =   6140
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Tetris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 80 Then MsgBox "Pause"
If KeyCode = 38 And Fig > 1 Then Rotera
If KeyCode = 37 Then
ok = 1
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True And X + i - 1 - 1 > 0 Then
If n(X + i - 1 - 1, Y - i2 + 1) = True Then ok = 0
End If
If s(i, i2) = True And X + i - 1 - 1 < 1 Then ok = 0
Next i2: Next i
If ok = 1 Then
X = X - 1
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True Then
If s(i + 1, i2) = False Then Shape1(coor(X + i, Y - i2 + 1)).FillColor = RGB(255, 255, 255)
End If
If s(i, i2) = True Then Shape1(coor(X + i - 1, Y - i2 + 1)).FillColor = Färg
Next i2: Next i
End If
End If
If KeyCode = 39 Then
ok = 1
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True And X + i - 1 + 1 < 11 Then
If n(X + i - 1 + 1, Y - i2 + 1) = True Then ok = 0
End If
If s(i, i2) = True And X + i - 1 + 1 > 10 Then ok = 0
Next i2: Next i
If ok = 1 Then
X = X + 1
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True Then
If s(i - 1, i2) = False Then Shape1(coor(X + i - 2, Y - i2 + 1)).FillColor = RGB(255, 255, 255)
End If
If s(i, i2) = True Then Shape1(coor(X + i - 1, Y - i2 + 1)).FillColor = Färg
Next i2: Next i
End If
End If
If KeyCode = 40 And Fig > 0 Then
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True Then Shape1(coor(X + i - 1, Y - i2 + 1)).FillColor = RGB(255, 255, 255)
Next i2: Next i
Do
Y = Y - 1
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True And Y - i2 + 1 = 1 Then Fig = 0
If Y - i2 > 0 And X + i - 1 > 0 And X + i - 1 < 11 Then
If s(i, i2) = True And n(X + i - 1, Y - i2) = True Then Fig = 0
End If
Next i2: Next i
Loop Until Fig = 0
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True Then Shape1(coor(X + i - 1, Y - i2 + 1)).FillColor = Färg
If s(i, i2) = True Then n(X + i - 1, Y - i2 + 1) = True
Next i2: Next i
End If
End Sub
Private Sub Form_Load()
For i = 1 To 4: For i2 = 1 To 4: s2(i, i2) = False: Next i2: Next i
For i = 1 To 10: For i2 = 1 To 25: n(i, i2) = False: Next i2: Next i
Score = -10
Timer1.Interval = Hast
For i = 0 To 11
Shape3(i).BorderColor = &HE0E0E0
Shape3(i).FillStyle = 0
Shape3(i).FillColor = &HE0E0E0
Next i
For i = 0 To 249
Shape1(i).BorderColor = RGB(255, 255, 255)
Shape1(i).FillStyle = 0
Shape1(i).FillColor = RGB(255, 255, 255)
Next i
Fig2 = 0
Nyfig
End Sub
Function coor(xx, yy)
coor = (yy - 1) * 10 + xx - 1
End Function
Private Sub Timer1_Timer()
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True And Y - i2 + 1 = 1 Then Fig = 0
If Y - i2 > 0 And X + i - 1 > 0 And X + i - 1 < 11 Then
If s(i, i2) = True And n(X + i - 1, Y - i2) = True Then Fig = 0
End If
Next i2: Next i
If Fig = 0 Then
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True Then n(X + i - 1, Y - i2 + 1) = True
Next i2: Next i
Ner
Nyfig
Else
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True And s(i, i2 - 1) = False Then Shape1(coor(X + i - 1, Y - i2 + 1)).FillColor = RGB(255, 255, 255)
Next i2: Next i
End If
Y = Y - 1
For i = 1 To 4: For i2 = 1 To 4
If s(i, i2) = True Then
If Shape1(coor(X + i - 1, Y - i2 + 1)).FillColor <> RGB(255, 255, 255) And Shape1(coor(X + i - 1, Y - i2 + 1)).FillColor <> Färg Then Gameover: Exit Sub
End If
If s(i, i2) = True Then Shape1(coor(X + i - 1, Y - i2 + 1)).FillColor = Färg
Next i2: Next i
End Sub
Public Sub Nyfig()
Timer1.Interval = Hast
Fig = Fig2
Färg = Färg2
Fig2 = Int(Rnd * 7) + 1
X = 4
Y = 26
Rot = 1
For i = 1 To 4: For i2 = 1 To 4: s(i, i2) = s2(i, i2): s2(i, i2) = 0: Next i2: Next i
Select Case Fig2
Case 1
s2(2, 2) = True
s2(3, 2) = True
s2(2, 3) = True
s2(3, 3) = True
Färg2 = RGB(255, 0, 0)
Case 2
s2(1, 2) = True
s2(2, 2) = True
s2(3, 2) = True
s2(4, 2) = True
Färg2 = RGB(0, 180, 0)
Case 3
s2(2, 1) = True
s2(3, 1) = True
s2(3, 2) = True
s2(3, 3) = True
Färg2 = RGB(255, 150, 0)
Case 4
s2(3, 1) = True
s2(2, 1) = True
s2(2, 2) = True
s2(2, 3) = True
Färg2 = RGB(100, 100, 100)
Case 5
s2(3, 1) = True
s2(3, 2) = True
s2(3, 3) = True
s2(2, 2) = True
Färg2 = RGB(200, 0, 200)
Case 6
s2(2, 1) = True
s2(2, 2) = True
s2(3, 2) = True
s2(3, 3) = True
Färg2 = RGB(100, 100, 255)
Case 7
s2(3, 1) = True
s2(3, 2) = True
s2(2, 2) = True
s2(2, 3) = True
Färg2 = RGB(0, 200, 200)
End Select
For i = 1 To 4: For i2 = 1 To 3
Shape3((i2 - 1) * 4 + i - 1).FillColor = &HE0E0E0
If s2(i, i2) = True Then Shape3((i2 - 1) * 4 + i - 1).FillColor = Färg2
Next i2: Next i
End Sub

Public Sub Rotera()
Rot2 = Rot + 1
If Rot2 = 5 Then Rot2 = 1
If (Fig = 2 Or Fig > 5) And Rot2 = 3 Then Rot2 = 1
For i = 1 To 4: For i2 = 1 To 4: s3(i, i2) = 0: Next i2: Next i
Select Case Fig
Case 2
Select Case Rot2
Case 1
s3(1, 2) = True
s3(2, 2) = True
s3(3, 2) = True
s3(4, 2) = True
Case 2
s3(2, 1) = True
s3(2, 2) = True
s3(2, 3) = True
s3(2, 4) = True
End Select
Case 3
Select Case Rot2
Case 1
s3(2, 1) = True
s3(3, 1) = True
s3(3, 2) = True
s3(3, 3) = True
Case 2
s3(4, 1) = True
s3(4, 2) = True
s3(3, 2) = True
s3(2, 2) = True
Case 3
s3(3, 3) = True
s3(2, 3) = True
s3(2, 2) = True
s3(2, 1) = True
Case 4
s3(2, 2) = True
s3(2, 1) = True
s3(3, 1) = True
s3(4, 1) = True
End Select
Case 4
Select Case Rot2
Case 1
s3(3, 1) = True
s3(2, 1) = True
s3(2, 2) = True
s3(2, 3) = True
Case 2
s3(2, 1) = True
s3(3, 1) = True
s3(4, 1) = True
s3(4, 2) = True
Case 3
s3(3, 1) = True
s3(3, 2) = True
s3(3, 3) = True
s3(2, 3) = True
Case 4
s3(2, 1) = True
s3(2, 2) = True
s3(3, 2) = True
s3(4, 2) = True
End Select
Case 5
Select Case Rot2
Case 1
s3(3, 1) = True
s3(3, 2) = True
s3(3, 3) = True
s3(2, 2) = True
Case 2
s3(3, 1) = True
s3(2, 2) = True
s3(3, 2) = True
s3(4, 2) = True
Case 3
s3(2, 1) = True
s3(2, 2) = True
s3(2, 3) = True
s3(3, 2) = True
Case 4
s3(2, 1) = True
s3(3, 1) = True
s3(4, 1) = True
s3(3, 2) = True
End Select
Case 6
Select Case Rot2
Case 1
s3(2, 1) = True
s3(2, 2) = True
s3(3, 2) = True
s3(3, 3) = True
Case 2
s3(2, 2) = True
s3(3, 2) = True
s3(3, 1) = True
s3(4, 1) = True
End Select
Case 7
Select Case Rot2
Case 1
s3(3, 1) = True
s3(3, 2) = True
s3(2, 2) = True
s3(2, 3) = True
Case 2
s3(2, 1) = True
s3(3, 1) = True
s3(3, 2) = True
s3(4, 2) = True
End Select
End Select
ok = 1
For i = 1 To 4: For i2 = 1 To 4
If s3(i, i2) = True Then
If X + i - 1 < 1 Or X + i - 1 > 10 Or Y - i2 + 1 < 1 Then ok = 0
If ok = 1 Then
If n(X + i - 1, Y - i2 + 1) = True Then ok = 0
End If
End If
Next i2: Next i
If ok = 0 Then Exit Sub
Rot = Rot2
For i = 1 To 4: For i2 = 1 To 4
If s3(i, i2) = True And s(i, i2) = False Then Shape1(coor(X + i - 1, Y - i2 + 1)).FillColor = Färg
If s3(i, i2) = False And s(i, i2) = True Then Shape1(coor(X + i - 1, Y - i2 + 1)).FillColor = RGB(255, 255, 255)
s(i, i2) = s3(i, i2)
Next i2: Next i
End Sub
Public Sub Ner()
Score = Score + 10
Label3.Caption = Score
For i2 = 25 To 1 Step -1
ok = 1
For i = 1 To 10
If n(i, i2) = False Then ok = 0
Next i
If ok = 1 Then

Tetris.SetFocus
Score = Score + 150
Label3.Caption = Score
For i = 1 To 10
For i3 = i2 To 24
n(i, i3) = n(i, i3 + 1)
Shape1(coor(i, i3)).FillColor = Shape1(coor(i, i3 + 1)).FillColor
Next i3
Next i
End If
Next i2
End Sub

Public Sub Gameover()

MsgBox "Your score was " & Score
Unload Me
Tetris2.Show
End Sub

Attribute VB_Name = "Module1"
Option Explicit

Public PozicijaLeft As Integer                             'left  da moze vratiti potez
Public PozicijaTop As Integer                              'top da moze vratiti potez
Public PozicijaLeft2 As Integer                            'left  da moze vratiti potez odnes.figura
Public PozicijaTop2 As Integer                             'top da moze vratiti potez odnes.figura
Public BrojOdneseneFigure As Integer
Public BrojFigure As Integer                               'broj indeksa pokrenute figure
Public bOO As Integer, wOO As Integer, bOOO As Integer, wOOO As Integer
Public Pomakni As Integer                                  'indeks figure koja je pomaknuta
Public Signal As Integer, IgraPrvi As Integer              'pok. kad je komj.na potezu,za ispis poteza
Public Baza(1 To 4) As Integer, Figure  As Integer
Public Sekunde1 As Long, Sekunde2 As Long, Vreme As Integer 'Broj sekundi levela
Public Ampasan As Integer


'Index de la piece qui bouge
Public Bouge As Integer
Public CaseFrom As String
Public CaseTo As String
Public LastKey As String
Public Piece As String


Public Plateau(64) As String
Public Moves

Public allMoves As New Collection

Public WhiteMove As Boolean
Public BlackMove As Boolean

Public GameSize As Double

Public cFile As New cOpen


Public FSO As New FileSystemObject
Public ts As TextStream

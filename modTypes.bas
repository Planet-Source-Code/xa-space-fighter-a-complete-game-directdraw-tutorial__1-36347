Attribute VB_Name = "modTypes"
'SpaceFighter v 1.00
'Type Definitions
Public Type Point
x As Long
y As Long
SpeedY As Long
End Type
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Type LocationSystem
x As Long
y As Long
Height As Long
Width As Long
SpeedX As Long
SpeedY As Long
End Type

Public Type AILogic
Location As LocationSystem
AIID As Long
Shooting As Boolean
PointsForDeath As Integer
LastShot As Long
Life As Integer
End Type

Public Declare Function GetTickCount Lib "kernel32" () As Long
'You might want to tamper around with those:
'but remember, nothing is good when it's too much
Public Const AllowYMovement As Boolean = False
Public Const AmountOfStars As Integer = 100
Public Const MaxStarSpeed As Integer = 5
Public Const MaxGoodBullets As Long = 500 'at a time
Public Const MaxBadBullets As Long = 1000 'at a time
'Those are just public arrays
Public Stars() As Point 'The stars behind us
Public GoodWeaponShots() As Point ' your laser
Public BadWeaponShots() As Point ' enemy laser

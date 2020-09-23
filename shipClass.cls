VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "shipClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************
'* Made by Ruble
'* ruble_19@yahoo.com
'**********************
'*
'* You may use this code any way you like.
'*
'* Description:
'* This is a shipClass that is used to store information
'* about each ship on the field. And has some functions
'* that pertain to the ship
'******************************************************************

Public headX As Integer     'X coord of ship start (0 - 10)
Public headY As Integer     'Y coord of ship start (0 - 10)
Public length As Integer    'Length of the ship
Public direction As Integer 'Direction of the ship (hor = vbKeyRight or vert = vbKeyDown)
Dim status() As Integer     'Stores info about each section of the ship (hit or not)
Sub drawShip()
'This sub draws the ship on field1 (left field)
'It draws the human opponents' ships

'The length determines how the ship is drawn
'I multiply each coord by 30 because thats the length
'of each section of the ship
Select Case length
Case 1
    BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipSingleHorMask.hDC, 0, 0, vbSrcAnd
    BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipSingleHor.hDC, 0, 0, vbSrcPaint
Case 2
    'Direction determines how the ship is drawn
    If direction = vbKeyRight Then
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEnd.hDC, 0, 0, vbSrcPaint
    ElseIf direction = vbKeyDown Then
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEnd.hDC, 0, 0, vbSrcPaint
    End If
Case 3
    'Direction determines how the ship is drawn
    If direction = vbKeyRight Then
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1Mask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, (headX + 2) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, (headX + 2) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEnd.hDC, 0, 0, vbSrcPaint
    ElseIf direction = vbKeyDown Then
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1DownMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1Down.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 2) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 2) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEnd.hDC, 0, 0, vbSrcPaint
    End If
Case 4
    'Direction determines how the ship is drawn
    If direction = vbKeyRight Then
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1Mask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, (headX + 2) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid2Mask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, (headX + 2) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid2.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, (headX + 3) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, (headX + 3) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEnd.hDC, 0, 0, vbSrcPaint
    ElseIf direction = vbKeyDown Then
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid2DownMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid2Down.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 2) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1DownMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 2) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1Down.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 3) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field1.hDC, headX * 30, (headY + 3) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEnd.hDC, 0, 0, vbSrcPaint
    End If
End Select
frmMain.field1.Refresh
End Sub
Sub drawShip2()
'This sub draws the ship on field2 (left field)
'It draws the computer generated ships

'The length determines how the ship is drawn
'I multiply each coord by 30 because thats the length
'of each section of the ship
Select Case length
Case 1
    BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipSingleHorMask.hDC, 0, 0, vbSrcAnd
    BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipSingleHor.hDC, 0, 0, vbSrcPaint
Case 2
    'Direction determines how the ship is drawn
    If direction = vbKeyRight Then
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEnd.hDC, 0, 0, vbSrcPaint
    ElseIf direction = vbKeyDown Then
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEnd.hDC, 0, 0, vbSrcPaint
    End If
Case 3
    'Direction determines how the ship is drawn
    If direction = vbKeyRight Then
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1Mask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, (headX + 2) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, (headX + 2) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEnd.hDC, 0, 0, vbSrcPaint
    ElseIf direction = vbKeyDown Then
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1DownMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1Down.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 2) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 2) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEnd.hDC, 0, 0, vbSrcPaint
    End If
Case 4
    'Direction determines how the ship is drawn
    If direction = vbKeyRight Then
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipLeftEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1Mask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, (headX + 1) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, (headX + 2) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid2Mask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, (headX + 2) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid2.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, (headX + 3) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, (headX + 3) * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipRightEnd.hDC, 0, 0, vbSrcPaint
    ElseIf direction = vbKeyDown Then
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, headY * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipUpEnd.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid2DownMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 1) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid2Down.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 2) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1DownMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 2) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipMid1Down.hDC, 0, 0, vbSrcPaint
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 3) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEndMask.hDC, 0, 0, vbSrcAnd
        BitBlt frmMain.field2.hDC, headX * 30, (headY + 3) * 30, frmMain.tile.ScaleWidth, frmMain.tile.ScaleHeight, frmMain.shipDownEnd.hDC, 0, 0, vbSrcPaint
    End If
End Select
frmMain.field2.Refresh
End Sub

Function sectionNumber(ByVal x, ByVal y) As Integer
'This function returns the section number of the ship
'from the start of it given x and y coord

Dim tempX As Integer 'temp X coord
Dim tempY As Integer 'temp Y coord

tempX = headX       'set temp X to head X of ship
tempY = headY       'set temp Y to head Y of ship
sectionNumber = 1   'initialize sectionNum to 1

'Direction determines how to look up the section
If direction = vbKeyRight Then
    'Loop through the ship's section until you hit the right one
    Do Until tempX = x
        tempX = tempX + 1
        sectionNumber = sectionNumber + 1
    Loop
Else
    'Loop through the ship's section until you hit the right one
    Do Until tempY = y
        tempY = tempY + 1
        sectionNumber = sectionNumber + 1
    Loop
End If
End Function

Sub setValues(ByVal x As Integer, ByVal y As Integer, ByVal s As Integer, ByVal dir As Integer)
'This initializes the ship

Dim i As Integer    'counter

headX = x   'head X coord of ship
headY = y   'head Y coord of ship
length = s  'size of ship
direction = dir 'direction of ship (vbKeyRight or vbKeyDown)

ReDim status(1 To s) 'Set the status array to be the size of ship

'Initialize status array to all 0's
For i = 1 To s
    status(i) = 0
Next i
End Sub
Function cripple() As Boolean
'This function updates the status of the ship
'when a section of it has been hit
'It returns false if the ship is crippled and
'true if the ship is sunk.

Dim i As Integer 'counter

'Basically everytime a ship gets hit I set the next
'available place in the status() to 1 meaning hit.
'When all of status() is equal to 1 then the ship
'is sunk

'Loop through array and set the next available space
'to 1
For i = 1 To length
    If status(i) = 0 Then
        status(i) = 1
        cripple = isDead() 'Checks if ship is sunk
        Exit Function
    End If
Next i
cripple = isDead() 'Checks if ship is sunk
End Function
Function isDead() As Boolean
'This basically just goes through the status array
'and if all of it is set to 1 then return true (for dead)

Dim i As Integer 'counter

For i = 1 To length
    If status(i) = 0 Then
        isDead = False
        Exit Function
    End If
Next i
isDead = True
End Function

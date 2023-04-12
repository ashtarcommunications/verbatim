VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCheatSheet 
   Caption         =   "Cheat Sheet"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmCheatSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCheatSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    Me.lboxShortcuts.AddItem "----------Speech----------"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Send To Speech"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "`/~ Key"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Send To Speech End"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Alt + `/~ Key"
    
    
    Me.lboxShortcuts.AddItem "----------Cells----------"
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Insert Cell Above"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "F3"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Insert Cell Below"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Alt + F3"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Merge Cells"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "F4"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Toggle Highlighting"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "F11"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Toggle Evidence"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "F7"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Toggle Group"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "F8"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Extend Argument"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "F9"
    
    
    Me.lboxShortcuts.AddItem "----------Rows----------"
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Insert Row Above"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "F5"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Insert Row Below"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Alt + F5"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Delete Row"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Ctrl + Alt + F5"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Move Selection Up"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Ctrl + Alt + Up"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Move Selection Down"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Ctrl + Alt + Down"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Go To Bottom"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Ctrl + Alt + Shift + Down"


    Me.lboxShortcuts.AddItem "----------Sheets----------"
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Add Aff Flow"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Ctrl + Alt + A"

    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Add Neg Flow"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Ctrl + Alt + N"

    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Add CX Flow"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Ctrl + Alt + X"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Next Flow"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Ctrl + PgUp"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Previous Flow"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "Ctrl + PgDown"
    
    
    Me.lboxShortcuts.AddItem "----------Insert----------"
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Enter Cell"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "F2"
    
    
    Me.lboxShortcuts.AddItem "----------Settings----------"
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Show Cheat Sheet"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = "F12"
End Sub


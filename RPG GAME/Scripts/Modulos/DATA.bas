Attribute VB_Name = "DATA"
Option Explicit

Public Type Sprite
            xCoord As Integer
            XPOS As Double
            yCoord As Integer
            YPOS As Double
            ID As String
        End Type
Public Type Position
            x As Double
            Y As Double
        End Type


Public Type Weapon
            NumID As Integer
            NumSlot As Integer
            ID As String
            Dmg As Integer
            Durabillity As Integer
            Weight As Integer
            Precision As Integer
        End Type
Public Type Item
            ID As String
            Qnt As Integer
            Durabillity As Integer
        End Type
Public Type Inventory
            InventoryName As String
            InventorySize As Integer
            ColumnID As Integer
            InventorySlots(120) As Item
        End Type
            

Public Type SceneType
            Layer1 As Worksheet
            Layer2 As Worksheet
            Layer3 As Worksheet
            XPOS As Integer
            YPOS As Integer
        End Type


Public Const xArraySize = 17
Public Const yArraySize = 17
Public Const zArraySize = 3
Public SpriteArray(xArraySize, yArraySize, zArraySize) As Sprite

Public HouseCount As Integer
Public TreeCount As Integer

Public ActualScene As SceneType

Public PlayerDirection As String 'TEXTURE
Public PlayerVar As clsPlayer

Public Const InventoryDisplaySize = 10
Public InventoryArray(999) As Inventory

Public VarQnt As Integer

Public Tree_Life As Integer
Public Max_Tree_life As Integer

'SCRIPT VARIABLES

Public ActualScript As Integer
Public ScriptCheck As Collection

'usfrmTalk_OptionMode
Public Talk_OptionNum As Integer
Public Talk_OptionTexts As Collection
Public Talk_SelectedOption As Integer


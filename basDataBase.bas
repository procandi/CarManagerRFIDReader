Attribute VB_Name = "basDataBase"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟資料庫有關的所有變數、常數、函式等的地方。                */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*msado15.dll。                                                   */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/02/26 */
'/******************************************************************/
Option Explicit


'/**************************公用的資料庫、資料表連結變數及物件***********************************/
Public Connection_String As String '最基本讓整個系統使用的資料庫連線字串
Public Connection As New ADODB.Connection '最基本讓整個系統使用的資料庫連線物件
Public Recordset As New ADODB.Recordset '最基本讓整個系統使用的資料表物件

Public Source_Connection_String As String '萬一以上的不夠用，另一個讓整個系統使用的資料庫連線字串
Public Source_Connection As New ADODB.Connection '萬一以上的不夠用，另一個讓讓整個系統使用的資料庫連線物件
Public Source_Recordset As New ADODB.Recordset '萬一以上的不夠用，另一個讓讓整個系統使用的資料表物件

Public Target_Connection_String As String '萬一以上的不夠用，另一個讓整個系統使用的資料庫連線字串
Public Target_Connection As New ADODB.Connection '萬一以上的不夠用，另一個讓讓整個系統使用的資料庫連線物件
Public Target_Recordset As New ADODB.Recordset '萬一以上的不夠用，另一個讓讓整個系統使用的資料表物件
'/**************************小華修改的(2009/02/25)***********************************/



'/*                 用於連線到SQLSERVER的函式，傳回真代表成功，假代表失敗                             */
'/*            Example Input Connect2SQLServer("127.0.0.1","DB1","User","Pass",Connection)            */
Public Function Connect2SQLServer(ByRef IP As String, ByRef DB As String, ByRef ID As String, ByRef PW As String, ByRef Connection As ADODB.Connection) As Boolean
    On Error GoTo err

    If Connection.State <> ADODB.adStateClosed Then
        Connection.Close
    End If
    
    Connection.Open "Provider=SQLOLEDB;Data Source=" & IP & ";Initial Catalog=" & DB & ";User Id=" & ID & ";Password=" & PW & ";"

    If Connection.State Then
        Connect2SQLServer = True
    Else
        Connect2SQLServer = False
    End If

    If False Then
err:
        Call ErrorOut("Function Connect2SQLServer Error!")
        Connect2SQLServer = False
    End If
End Function
'/**************************小華修改的(2009/02/03)***********************************/



'/*                     用於連線到各種資料庫的函式，傳回真代表成功，假代表失敗                    */
'/*            Example Input Connect2DataBase("Server=127.0.0.1;DRIVER={SQL Server};UID=User;PWD=Pass;DATABASE=DB1;",,Connection)            */
Public Function Connect2DataBase(ByRef Connection_String As String, ByRef Connection As ADODB.Connection) As Boolean
    On Error GoTo err

    If Connection.State <> ADODB.adStateClosed Then
        Connection.Close
    End If
    
    Connection.Open Connection_String

    If Connection.State Then
        Connect2DataBase = True
    Else
        Connect2DataBase = False
    End If

    If False Then
err:
        Call ErrorOut("Function Connect2DataBase Error!")
        Connect2DataBase = False
    End If
End Function
'/**************************小華修改的(2009/02/03)***********************************/




'/*                      開啟資料表的函式，傳回真代表成功，假代表失敗                */
'/*            Example Input OpenRecordset("select * from NorthWind",Connection,Recordset)            */
Public Function OpenRecordset(ByRef SQL As String, ByRef Connection As ADODB.Connection, ByRef Recordset As ADODB.Recordset, Optional ByVal CursorType As Integer, Optional ByVal LockType As Integer) As Boolean
    On Error GoTo err
    
    If Recordset.State <> ADODB.adStateClosed Then
        Recordset.Close
    End If
     
    If CursorType = 0 Then
        CursorType = ADODB.adOpenStatic
    End If
    If LockType = 0 Then
        LockType = ADODB.adLockOptimistic
    End If
    
    Recordset.Open SQL, Connection, CursorType, LockType
    
    If Recordset.State Then
        OpenRecordset = True
    Else
        OpenRecordset = False
    End If

    If False Then
err:
        Call ErrorOut("Function OpenRecordset Error!")
        OpenRecordset = False
    End If
End Function
'/**************************小華修改的(2009/02/04)***********************************/



'/*    主要用途為用於資料庫的where判斷部份，能因傳入的欄位是不是null而修改傳回的sql指令           */
'/*               Example Input OpenRecordset(null)            */
Public Function isFieldsNull(ByRef Fields As String) As String
    If Fields = "Null" Then
        isFieldsNull = " is null"
    Else
        isFieldsNull = "='" & Fields & "'"
    End If
End Function
'/**************************小華修改的(2009/02/25)***********************************/


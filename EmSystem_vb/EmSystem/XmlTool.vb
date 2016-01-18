Imports System.Xml

Public Class XmlTool
    Public Class XmlObject
        ''' <summary>
        ''' 第幾層
        ''' </summary>
        ''' <returns></returns>
        Public Property level As Integer
        ''' <summary>
        ''' tag名稱
        ''' </summary>
        ''' <returns></returns>
        Public Property name As String
        ''' <summary>
        ''' 上層tag名稱
        ''' </summary>
        ''' <returns></returns>
        Public Property upName As String
        ''' <summary>
        ''' InnerText,夾在tag中間的字串
        ''' </summary>
        ''' <returns></returns>
        Public Property text As String
        ''' <summary>
        ''' 值(不用)
        ''' </summary>
        ''' <returns></returns>
        Public Property value As String
    End Class

    ''' <summary>
    ''' 建議讀取方式，1:不重複tag名稱的資料，2:具有重複Tag名稱的資料
    ''' </summary>
    Public loadType As String = "1"
    Private objList As List(Of XmlObject)

    ''' <summary>
    ''' 把路徑或URL中的XML，轉換成XmlObject的List
    ''' </summary>
    ''' <param name="path">檔案路徑或URL、介接可用</param>
    ''' <returns></returns>
    Public Function Load(path As String) As List(Of XmlObject)
        Dim myXD As XmlDocument = New XmlDocument()
        myXD.Load(path)
        objList = New List(Of XmlObject)
        LoadXmlNode(myXD.ChildNodes, 0, "")
        Return objList
    End Function

    ''' <summary>
    ''' 取的某個上層名稱的集合
    ''' </summary>
    ''' <param name="objList">使用Load出來的結果，或其他List XmlObject</param>
    ''' <param name="_upName">上層Tag的名稱</param>
    ''' <returns></returns>
    Public Function GatObjArray(objList As List(Of XmlObject), _upName As String) As XmlObject()
        Return (From obj In objList Where obj.upName = _upName Select obj).ToArray()
    End Function

    ''' <summary>
    ''' 取的某個某層的集合
    ''' </summary>
    ''' <param name="objList">使用Load出來的結果，或其他List XmlObject</param>
    ''' <param name="_level">第幾層</param>
    ''' <returns></returns>
    Public Function GatObjArray(objList As List(Of XmlObject), _level As Integer) As XmlObject()
        Return (From obj In objList Where obj.level = _level Select obj).ToArray()
    End Function

    ''' <summary>
    ''' 取的某層且某個上層名稱的集合
    ''' </summary>
    ''' <param name="objList">使用Load出來的結果，或其他List XmlObject</param>
    ''' <param name="_upName">上層Tag的名稱</param>
    ''' <param name="_level">第幾層</param>
    ''' <returns></returns>
    Public Function GatObjArray(objList As List(Of XmlObject),_upName As string, _level As Integer) As XmlObject()
        Return (From obj In objList Where obj.level = _level And obj.upName = _upName Select obj).ToArray()
    End Function

    Private Sub LoadXmlNode(itemList As XmlNodeList, level As Integer, upName As String)
        For Each item As XmlNode In itemList
            If item.ChildNodes.Count > 0 Then
                If item.ChildNodes.Item(0).HasChildNodes Then
                    objList.Add(addXmlObj(level, item.Name, upName, "", item.Value))
                    LoadXmlNode(item.ChildNodes, level + 1, item.Name)
                Else
                    objList.Add(addXmlObj(level, item.Name, upName, item.InnerText, item.Value))
                    If loadType = "2" Then
                        LoadXmlNode(item.ChildNodes, level + 1, item.Name)
                    End If
                End If
            Else
                objList.Add(addXmlObj(level, item.Name, upName, item.InnerText, item.Value))
            End If
        Next
    End Sub

    Private Function addXmlObj(level As Integer, name As String,upName As String, text As String, Optional value As String = "") As XmlObject
        Dim obj As New XmlObject
        obj.level = level
        obj.name = name
        obj.upName = upName
        obj.text = text
        obj.value = value
        Return obj
    End Function
End Class

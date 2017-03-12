# Google_Map_Data
Get Google Map Data from Excel VBA 
...

应用案例：
  - 给予起点国家（例如：New York City,NY) 和目的地国家（例如：Chicago,IL) 算出两点之间以car为交通工具的时间(time)和距离(distance)
  - 给予一个国家，返回该国家的地理位置（经度和纬度）
  
   #### Google API to get distance & duration time
   ```vb
   https://maps.googleapis.com/maps/api/distancematrix/json?origins=Vancouver+BC|Seattle&destinations=San+Francisco|Victoria+BC&mode=bicycling&language=en
   ```
   #### Google API to get latitude &longitude
   ```vb
   http://maps.google.cn/maps/api/geocode/json?address=地址
   ```
<!--more-->
 代码1：获得国家之间的地理位置距离（by car)
 ```vb
 Public Function distanceDataFun(start As String, dest As String)
     Dim firstVal As String, secondVal As String, lastVal As String
     firstVal = "http://maps.googleapis.com/maps/api/distancematrix/json?origins="
     secondVal = "&destinations="
     lastVal = "&mode=car&language=pl&sensor=false"
     Set objHTTP = CreateObject("Msxml2.XMLHTTP.3.0")
     URL = firstVal & Replace(start, " ", "+") & secondVal & Replace(dest, " ", "+") & lastVal
     objHTTP.Open "GET", URL, False
     objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
     objHTTP.send ("")
     If InStr(objHTTP.responseText, """distance"" : {") = 0 Then GoTo ErrorHandl
     Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = "distance(?:.|\n)*?""value"".*?([0-9]+)": regex.Global = True
     Set matches = regex.Execute(objHTTP.responseText)
     distanceDataFun = Format(matches(0).SubMatches(0) / 1000) + " km"
     Exit Function
 ErrorHandl:
     distanceDataFun = -1
 End Function
 ```
 
  代码2：获得国家之间的时间段
  ```vb
Public Function durationDataFun(start As String, dest As String)
    Dim firstVal As String, secondVal As String, lastVal As String
    firstVal = "http://maps.googleapis.com/maps/api/distancematrix/json?origins="
    secondVal = "&destinations="
    lastVal = "&mode=car&language=pl&sensor=false"
    Set objHTTP = CreateObject("Msxml2.XMLHTTP.3.0")
    URL = firstVal & Replace(start, " ", "+") & secondVal & Replace(dest, " ", "+") & lastVal
    objHTTP.Open "GET", URL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    If InStr(objHTTP.responseText, """duration"" : {") = 0 Then GoTo ErrorHandl
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = "duration(?:.|\n)*?""value"".*?([0-9]+)": regex.Global = True
    Set matches = regex.Execute(objHTTP.responseText)
    Dim result As String
    result = ""
    If VBA.Int(matches(0).SubMatches(0) / 60 / 60 / 24) > 0 Then
        result = result + Format(matches(0).SubMatches(0) / 60 / 60, ".00") + " hour "
        'result = result + Format(VBA.Int(matches(0).SubMatches(0) / 60 / 60 / 24)) + " day " + Format(Int(matches(0).SubMatches(0) / 60 / 60) Mod 24) + " hour " + Format(Int(matches(0).SubMatches(0) / 60) Mod 60) + " min"
    Else
        'result = result + Format(Int(matches(0).SubMatches(0) / 60 / 60) Mod 24) + " hour " + Format((matches(0).SubMatches(0) / 60) Mod 60) + " min"
        result = result + Format(matches(0).SubMatches(0) / 60 / 60, ".00") + " hour "
    End If
    durationDataFun = result
    Exit Function
ErrorHandl:
    durationDataFun = -1
End Function
  ```
  
  代码3：获得国家的坐标
  ```vb
  Public Function latDataFun(val As String)
      Set objHTTP = CreateObject("Msxml2.XMLHTTP.3.0")
      URL = "http://maps.google.cn/maps/api/geocode/json?&address=" + val
      objHTTP.Open "GET", URL, False
      objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
      objHTTP.send ("")
      If InStr(objHTTP.responseText, """location"" : {") = 0 Then GoTo ErrorHandl
      Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = "location(?:.|\n)*?""lat"".*?([0-9])+\.([0-9]+)": regex.Global = True
      Set matches = regex.Execute(objHTTP.responseText)
      regex.Pattern = "([0-9])+\.([0-9]+)":
      Set Results = regex.Execute(matches(0))
      latDataFun = Results(0)
      Exit Function
  ErrorHandl:
      latDataFun = -1
  End Function
  Public Function lngDataFun(val As String)
      Set objHTTP = CreateObject("Msxml2.XMLHTTP.3.0")
      URL = "http://maps.google.cn/maps/api/geocode/json?&address=" + val
      objHTTP.Open "GET", URL, False
      objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
      objHTTP.send ("")
      If InStr(objHTTP.responseText, """location"" : {") = 0 Then GoTo ErrorHandl
      Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = "location(?:.|\n)*?""lng"".*?([0-9])+\.([0-9]+)": regex.Global = True
      Set matches = regex.Execute(objHTTP.responseText)
      regex.Pattern = "lng.*?([0-9])+\.([0-9]+)":
      Set matches2 = regex.Execute(matches(0))
      regex.Pattern = ".?([0-9])+\.([0-9]+)":
      Set result = regex.Execute(matches2(0))
      lngDataFun = result(0)
      Exit Function
  ErrorHandl:
      lngDataFun = -1
  End Function

  ```

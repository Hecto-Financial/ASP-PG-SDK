<!--#include virtual="/npg/inc/json2.asp"-->
<%


'로그 메시지 출력(파일)
Sub log_message(log_fname, msg)
    Dim fs, fname, dir
    Dim now_micro : now_micro  = right("000" & (timer * 1000) Mod 1000, 3) '밀리초
    Dim filename : filename = log_fname & "." & formatdatetime( date, 2 ) & ".log" '로그파일명 패턴
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' LOG_DIR(config.asp에서 설정)가 존재할 경우에만 파일 생성    
    If fs.FolderExists(LOG_DIR) = True Then
        Set dir = fs.getFolder(LOG_DIR)
        ' 파일 오픈( 없으면 생성 )
        Set fname = fs.OpenTextFile( dir & "/" & filename, 8, True )

        fname.WriteLine( formatdatetime(now , 0) & "." & now_micro & " : " & msg) '로그 메세지 패턴
        fname.Close

        Set dir = Nothing
        Set fname = Nothing
    End if
    Set fs=nothing
End Sub


'HTTP Post
'@param url : API URL
'@param jsonData : JSON request data 
'@param connTimeout : resolve, connect, send timeout
'@param readTimeout : receive timeout
'@return responseString : JSON from settlebank
Function sendPost( url, jsonData, connTimeout, readTimeout)
    Dim responseString
    
    Dim http : set http = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    call http.Open("POST", url, False)
    call http.SetTimeouts(connTimeout, connTimeout, connTimeout, readTimeout)
    call http.SetRequestHeader("Content-Type", "application/json;charset=UTF-8")
    call http.send(jsonData)

    if http.Status = 200 then
        Set responseString = JSON.parse(http.responseText)
    else
        Set responseString = JSON.parse("{ ""params"" : {}, ""data"":{} }")
        responseString.params.set "outStatCd", "0099"
        responseString.params.set "outRsltMsg", http.Status & " : " & http.StatusText
    end If

    Set sendPost = responseString
    Set responseString = Nothing
    Set http = Nothing
End Function



%>
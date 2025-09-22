<%
    
    '노티 출력
    Function printNoti(noti)
        Dim rslt : set rslt = noti
        Dim logStr : logStr = ""
        Dim key, value, idx
        key = noti.Keys
        value = noti.Items
        For idx = 0 To noti.Count -1
            logStr = logStr & key(idx) & "(" & value(idx) & ") "
        Next

        printNoti = logStr
        set rslt = nothing
    End Function
    
    '노티를 성공적으로 수신한 경우 처리할 로직을 작성하여 주세요.
    Function noti_success(noti)
        'TODO : 관련 로직 추가
        noti_success =  true
    End Function
    
    '입금대기시 처리할 로직을 작성하여 주세요.
    Function noti_waiting_pay(noti)
        'TODO : 관련 로직 추가
        noti_waiting_pay = true
    End Function
    
    '노티 수신중 해시 체크 에러가 생긴 경우 처리할 로직을 작성하여 주세요.
    Function noti_hash_error(noti)
        'TODO : 관련 로직 추가
        noti_hash_error = false
    End Function
%>
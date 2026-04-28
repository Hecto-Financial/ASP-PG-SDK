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
    '결제 완료(outStatCd=0021) 시 호출됩니다.
    '예시) 주문 상태를 "결제 완료"로 변경, 상품/서비스 제공, 고객에게 결제 완료 안내 등
    Function noti_success(noti)
        'TODO : 아래 로직을 가맹점 환경에 맞게 구현하세요.
        '  - DB에서 mchtTrdNo(상점주문번호)로 주문을 조회하여 상태 업데이트
        '  - 중복 처리 방지: 이미 처리된 주문인지 확인 후 처리
        '  - 예) DB Update: mchtTrdNo = noti.Item("mchtTrdNo"), trdAmt = noti.Item("trdAmt")
        noti_success =  true
    End Function

    '입금대기시 처리할 로직을 작성하여 주세요.
    '가상계좌 채번 완료(outStatCd=0051) 시 호출됩니다.
    '예시) 주문 상태를 "입금 대기"로 변경, 가상계좌번호/입금기한 저장, 고객에게 계좌 안내 등
    Function noti_waiting_pay(noti)
        'TODO : 아래 로직을 가맹점 환경에 맞게 구현하세요.
        '  - DB에서 mchtTrdNo(상점주문번호)로 주문 상태를 "입금 대기"로 업데이트
        '  - vAcntNo(가상계좌번호), expireDt(입금만료일시) 저장
        '  - 아직 결제가 완료된 것이 아니므로 상품/서비스 제공 금지
        noti_waiting_pay = true
    End Function

    '노티 수신중 해시 체크 에러가 생긴 경우 처리할 로직을 작성하여 주세요.
    '해시 불일치는 데이터 위변조 가능성이 있으므로 반드시 원인을 파악하고 대처해야 합니다.
    Function noti_hash_error(noti)
        'TODO : 아래 로직을 가맹점 환경에 맞게 구현하세요.
        '  - 관리자에게 해시 오류 알림 발송 (이메일, 슬랙 등)
        '  - 오류 로그 기록 (mchtTrdNo, 수신된 pktHash 등)
        noti_hash_error = false
    End Function
%>
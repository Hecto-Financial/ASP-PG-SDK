<!--#include virtual="/npg/inc/config.asp"-->
<!--#include virtual="/npg/inc/settleUtils.asp"-->
<!--#include virtual="/npg/inc/KISA_SHA256.asp"-->
<!--#include virtual="/npg/processNoti.asp"-->
<%

'이 페이지는 수정시 주의가 필요합니다. 수정시 html태그나 자바스크립트가 들어가는 경우 동작을 보장할 수 없습니다

'============================================================================================================================================
'   노티 파라미터 수신( 헥토파이낸셜 -> 가맹점 )
'============================================================================================================================================
Dim noti : Set noti = server.CreateObject("Scripting.Dictionary")
noti.Add "outStatCd",       "" & request.Form("outStatCd")      '거래상태
noti.Add "trdNo",           "" & request.Form("trdNo")          '거래번호
noti.Add "method",          "" & request.Form("method")         '결제수단
noti.Add "bizType",         "" & request.Form("bizType")        '업무구분
noti.Add "mchtId",          "" & request.Form("mchtId")         '상점아이디
noti.Add "mchtTrdNo",       "" & request.Form("mchtTrdNo")      '상점주문번호
noti.Add "mchtCustNm",      "" & request.Form("mchtCustNm")     '고객명
noti.Add "mchtName",        "" & request.Form("mchtName")       '상점한글명
noti.Add "pmtprdNm",        "" & request.Form("pmtprdNm")       '상품명
noti.Add "trdDtm",          "" & request.Form("trdDtm")         '거래일시
noti.Add "trdAmt",          "" & request.Form("trdAmt")         '거래금액
noti.Add "billKey",         "" & request.Form("billKey")        '자동결제키
noti.Add "billKeyExpireDt", "" & request.Form("billKeyExpireDt")'자동결제키 유효기간
noti.Add "bankCd",          "" & request.Form("bankCd")         '은행코드
noti.Add "bankNm",          "" & request.Form("bankNm")         '은행명
noti.Add "cardCd",          "" & request.Form("cardCd")         '카드사코드
noti.Add "cardNm",          "" & request.Form("cardNm")         '카드명
noti.Add "telecomCd",       "" & request.Form("telecomCd")      '이통사코드
noti.Add "telecomNm",       "" & request.Form("telecomNm")      '이통사명
noti.Add "vAcntNo",         "" & request.Form("vAcntNo")        '가상계좌번호
noti.Add "expireDt",        "" & request.Form("expireDt")       '가상계좌 입금만료일시
noti.Add "AcntPrintNm",     "" & request.Form("AcntPrintNm")    '통장인자명
noti.Add "dpstrNm",         "" & request.Form("dpstrNm")        '입금자명
noti.Add "email",           "" & request.Form("email")          '고객이메일
noti.Add "mchtCustId",      "" & request.Form("mchtCustId")     '상점고객아이디
noti.Add "cardNo",          "" & request.Form("cardNo")         '카드번호
noti.Add "cardApprNo",      "" & request.Form("cardApprNo")     '카드승인번호
noti.Add "instmtMon",       "" & request.Form("instmtMon")      '할부개월수
noti.Add "instmtType",      "" & request.Form("instmtType")     '할부타입
noti.Add "phoneNoEnc",      "" & request.Form("phoneNoEnc")     '휴대폰번호(암호화)
noti.Add "orgTrdNo",        "" & request.Form("orgTrdNo")       '원거래번호
noti.Add "orgTrdDt",        "" & request.Form("orgTrdDt")       '원거래일자
noti.Add "mixTrdNo",        "" & request.Form("mixTrdNo")       '복합결제 거래번호
noti.Add "mixTrdAmt",       "" & request.Form("mixTrdAmt")      '복합결제 금액
noti.Add "payAmt",          "" & request.Form("payAmt")         '실결제금액
noti.Add "csrcIssNo",       "" & request.Form("csrcIssNo")      '현금영수증 승인번호
noti.Add "cnclType",        "" & request.Form("cnclType")       '취소거래타입
noti.Add "mchtParam",       "" & request.Form("mchtParam")      '기타주문정보
noti.Add "pktHash",         "" & request.Form("pktHash")        '해쉬값
noti.Add "acntType",         "" & request.Form("acnType")        '계좌구분
noti.Add "kkmAmt",         "" & request.Form("kkmAmt")        '카카오머니 금액
noti.Add "coupAmt",         "" & request.Form("coupAmt")        '쿠폰금액


call log_message(NOTI_LOG_FILE, "[" & noti.Item("mchtTrdNo") & "]params:" & printNoti(noti) )

'============================================================================================================================================
'   SHA256 해쉬 처리
'조합 필드 : 결과코드 + 거래일시 + 상점아이디 + 가맹점거래번호 + 거래금액(평문) + 라이센스키
'============================================================================================================================================
Dim hashPlain : hashPlain = ""
Dim hashCipher : hashCipher = ""
hashPlain = noti.Item("outStatCd") &_
            noti.Item("trdDtm") &_
            noti.Item("mchtId") &_
            noti.Item("mchtTrdNo") &_
            noti.Item("trdAmt") &_
            LICENSE_KEY
On error resume next
hashCipher = SHA256_Encrypt(hashPlain)'해쉬 값 계산
If Err.Number <> 0 Then
    call log_message(NOTI_LOG_FILE, "[" & noti.Item("mchtTrdNo") & "][SHA256 HASHING] Hashing Fail! " & err.number & " : " & err.description  )
else
    call log_message(NOTI_LOG_FILE, "[" & noti.Item("mchtTrdNo") & "][SHA256 HASHING] Plain Text[" & hashPlain & "] ---> Cipher Text[" & hashCipher & "]")
End If
On Error GoTo 0

'============================================================================================================================================
'   데이터 위변조 체크(해쉬 체크)
'hash데이타값이 맞는 지 확인 하는 루틴은 헥토파이낸셜에서 받은 데이타가 맞는지 확인하는 것이므로 꼭 사용하셔야 합니다
'정상적인 결제 건임에도 불구하고 노티 페이지의 오류나 네트웍 문제 등으로 인한 hash 값의 오류가 발생할 수도 있습니다.
'그러므로 hash 오류건에 대해서는 오류 발생시 원인을 파악하여 즉시 수정 및 대처해 주셔야 합니다. 
'그리고 정상적으로 데이터를 처리한 경우에도 헥토파이낸셜에서 응답을 받지 못한 경우는 결제결과가 중복해서 나갈 수 있으므로 관련한 처리도 고려되어야 합니다
'============================================================================================================================================
Dim resp : resp = false '노티 처리 결과
if hashCipher = noti.Item("pktHash") then
    call log_message(NOTI_LOG_FILE, "[" & noti.Item("mchtTrdNo") & "][SHA256 Hash Check] hashCipher[" & hashCipher & "] pktHash[" & noti.Item("pktHash") & "] equals?[TRUE]")
    if "0021" = noti.Item("outStatCd") then
        resp = noti_success(noti)
    elseif "0051" = noti.Item("outStatCd") then
        resp = noti_Waiting_pay(noti)
    else
        call log_message(NOTI_LOG_FILE, "[" & noti.Item("mchtTrdNo") & "][Undefined Code] outStatCd:" & noti.Item("outStatCd"))
        resp = false
    end if
else 
    call log_message(NOTI_LOG_FILE, "[" & noti.Item("mchtTrdNo") & "][SHA256 Hash Check] hashCipher[" & hashCipher & "] pktHash[" & noti.Item("pktHash") & "] equals?[FALSE]")
    resp = noti_hash_error(noti)
end if
'============================================================================================================================================
'   OK, FAIL 문자열은 헥토파이낸셜로 전송되어야 하는 값이므로 변경하거나 삭제하지마십시오.
'============================================================================================================================================
if resp then
    response.write("OK")
    call log_message(NOTI_LOG_FILE, "[" & noti.Item("mchtTrdNo") & "][Result] OK")
else
    response.write("FAIL")
    call log_message(NOTI_LOG_FILE, "[" & noti.Item("mchtTrdNo") & "][Result] FAIL")
end if

set noti = nothing
%>





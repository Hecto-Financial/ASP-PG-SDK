<% @CODEPAGE="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset = "UTF-8" %>
<!--#include virtual="/npg/inc/KISA_SHA256.asp"-->
<!--#include virtual="/npg/inc/settleUtils.asp"-->
<!--#include virtual="/npg/inc/config.asp"-->
<%
'============================================================================================================================================
'   신용카드 빌키 삭제 API 결과 처리
'   - 호출 경로: /spay/APICardActionDelkey.do (게이트웨이, JSON, UI 없음 / 서버 간 통신)
'   - 발급받은 빌키를 삭제(해지)합니다. 빌키는 평문 전송이며 암호화 대상이 아닙니다.
'============================================================================================================================================

'============================================================================================================================================
'   요청 파라미터
'============================================================================================================================================
Dim REQ_PARAMS : set REQ_PARAMS = JSON.parse("{ ""params"":{}, ""data"":{} }")
REQ_PARAMS.params.set "mchtId",     "" & request.Form("mchtId")     '상점아이디
REQ_PARAMS.params.set "ver",        "" & request.Form("ver")        '버전
REQ_PARAMS.params.set "method",     "" & request.Form("method")     '결제수단(CA)
REQ_PARAMS.params.set "bizType",    "" & request.Form("bizType")    '업무구분(A1:빌키 삭제)
REQ_PARAMS.params.set "encCd",      "" & request.Form("encCd")      '암호화구분
REQ_PARAMS.params.set "mchtTrdNo",  "" & request.Form("mchtTrdNo")  '상점주문번호
REQ_PARAMS.params.set "trdDt",      "" & request.Form("trdDt")      '요청일자
REQ_PARAMS.params.set "trdTm",      "" & request.Form("trdTm")      '요청시간
REQ_PARAMS.params.set "mobileYn",   "" & request.Form("mobileYn")   '모바일여부
REQ_PARAMS.params.set "osType",     "" & request.Form("osType")     '운영체제 구분

'요청 파라미터(바디) - 빌키(평문 전송, 암호화 대상 아님) + 해지사유코드(선택)
REQ_PARAMS.data.set "pktHash",      ""                              '해쉬값
REQ_PARAMS.data.set "billKey",      "" & request.Form("billKey")    '삭제할 빌키
REQ_PARAMS.data.set "etcInfo",      "" & request.Form("etcInfo")    '해지사유코드(선택)


'============================================================================================================================================
'   응답 파라미터 선언
'============================================================================================================================================
Dim RES_PARAMS : set RES_PARAMS = JSON.parse("{ ""params"":{}, ""data"":{} }")
RES_PARAMS.params.set "mchtId", ""      '상점아이디
RES_PARAMS.params.set "ver", ""         '버전
RES_PARAMS.params.set "method", ""      '결제수단
RES_PARAMS.params.set "bizType", ""     '업무구분
RES_PARAMS.params.set "encCd", ""       '암호화구분
RES_PARAMS.params.set "mchtTrdNo", ""   '상점주문번호
RES_PARAMS.params.set "trdNo", ""       '헥토파이낸셜거래번호
RES_PARAMS.params.set "trdDt", ""       '거래일자
RES_PARAMS.params.set "trdTm", ""       '거래시간
RES_PARAMS.params.set "outStatCd", ""   '결과코드
RES_PARAMS.params.set "outRsltCd", ""   '거절코드
RES_PARAMS.params.set "outRsltMsg", ""  '결과메세지

RES_PARAMS.data.set "pktHash", ""       '해쉬값
RES_PARAMS.data.set "billKey", ""       '삭제된 빌키


'============================================================================================================================================
'   SHA256 해쉬 처리
'조합 필드 : 요청일자 + 요청시간 + 상점아이디 + 상점주문번호 + "0" + 라이센스키
'※ 빌키 삭제는 거래금액이 없으므로 금액 자리에 반드시 "0"(문자)을 사용합니다.
'============================================================================================================================================
Dim hashPlain : hashPlain = REQ_PARAMS.params.get("trdDt") & REQ_PARAMS.params.get("trdTm") & REQ_PARAMS.params.get("mchtId") & REQ_PARAMS.params.get("mchtTrdNo") & "0" & LICENSE_KEY
Dim hashCipher : hashCipher = ""

On error resume next
hashCipher = SHA256_Encrypt(hashPlain)'해쉬 값 계산
If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SHA256 HASHING] Hashing Fail! " & err.number & " : " & err.description  )
else
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SHA256 HASHING] Plain Text[" & hashPlain & "] ---> Cipher Text[" & hashCipher & "]")
    REQ_PARAMS.data.set "pktHash", hashCipher 'SHA256 해쉬 결과 저장
End If
On Error GoTo 0


'============================================================================================================================================
'   AES256 암호화 처리 - 빌키 삭제는 암호화 대상 파라미터가 없습니다.
'============================================================================================================================================


'============================================================================================================================================
'   타겟 URL 설정
'  타겟 서버 : (tb)gw.settlebank.co.kr
'  신용카드 빌키 삭제 : ~/spay/APICardActionDelkey.do
'============================================================================================================================================
Dim requestUrl : requestUrl = CANCEL_SERVER & "/spay/APICardActionDelkey.do"
call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SEND POST] Request url :  " & requestUrl  )


'============================================================================================================================================
'   API호출(가맹점->세틀) 및 응답 처리
'============================================================================================================================================
Dim resData
On error resume next
call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SEND POST] Request :  " & JSON.stringify(REQ_PARAMS)  )
set resData = sendPost(requestUrl, JSON.stringify(REQ_PARAMS) , CONN_TIMEOUT, READ_TIMEOUT)
If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SEND POST] Fail! " & err.number & " : " & err.description  )
Else
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SEND POST] Response : " &  JSON.stringify(resData) )
End If
On Error GoTo 0


'============================================================================================================================================
'   응답 파라미터 세팅
'============================================================================================================================================
if resData.get("params") <> "" then
    RES_PARAMS.params.set "mchtId", resData.params.get("mchtId")            '상점아이디
    RES_PARAMS.params.set "ver", resData.params.get("ver")                  '버전
    RES_PARAMS.params.set "method", resData.params.get("method")            '결제수단
    RES_PARAMS.params.set "bizType", resData.params.get("bizType")          '업무구분
    RES_PARAMS.params.set "encCd", resData.params.get("encCd")              '암호화구분
    RES_PARAMS.params.set "mchtTrdNo", resData.params.get("mchtTrdNo")      '상점주문번호
    RES_PARAMS.params.set "trdNo", resData.params.get("trdNo")              '헥토파이낸셜거래번호
    RES_PARAMS.params.set "trdDt", resData.params.get("trdDt")              '거래일자
    RES_PARAMS.params.set "trdTm", resData.params.get("trdTm")              '거래시간
    RES_PARAMS.params.set "outStatCd", resData.params.get("outStatCd")      '결과코드
    RES_PARAMS.params.set "outRsltCd", resData.params.get("outRsltCd")      '거절코드
    RES_PARAMS.params.set "outRsltMsg", resData.params.get("outRsltMsg")    '결과메세지
end if
if resData.get("data") <> "" then
    RES_PARAMS.data.set "pktHash", resData.data.get("pktHash")              '해쉬값
    RES_PARAMS.data.set "billKey", resData.data.get("billKey")             '삭제된 빌키
end if


'============================================================================================================================================
'   AES256 복호화 처리 - 빌키 삭제는 복호화 대상 파라미터가 없습니다.
'============================================================================================================================================

%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="ie=edge">
<title>빌키 삭제 응답 결과</title>
<style type="text/css">
    body            {font-family:굴림; font-size:10pt; color:#000000;}
    td              {font-family:굴림; font-size:10pt; padding:3px; border:1px solid #e1e1e1;}
    table           {width:100%; border-collapse:collapse;}
    .left           {padding-left:5px; width:230px;}
    .right          {padding-left:5px;}
    .wrapper        {width:700px; border:1px solid #e1e1e1;}
    .tab            {background-color:#f1f1f1; padding:10px 20px; border:1px solid #e1e1e1; font-weight:bold; font-size:1.1em;}
    .button         {padding:5px 20px; border-radius:20px; border:1px solid #ccc; width:70%; margin:5px 0px; transition:0.3s; cursor:pointer;}
    .button:hover   {background-color:#aaaaaa;}
</style>
</head>
<body>
<h2>빌키 삭제 응답 결과</h2>
<div class="wrapper">
    <div class="tab">params</div>
    <table>
        <tr><td class="left">mchtId[상점아이디]</td><td class="right"><%= RES_PARAMS.params.get("mchtId") %></td></tr>
        <tr><td class="left">ver[버전]</td><td class="right"><%= RES_PARAMS.params.get("ver") %></td></tr>
        <tr><td class="left">method[결제수단]</td><td class="right"><%= RES_PARAMS.params.get("method") %></td></tr>
        <tr><td class="left">bizType[업무구분]</td><td class="right"><%= RES_PARAMS.params.get("bizType") %></td></tr>
        <tr><td class="left">encCd[암호화구분]</td><td class="right"><%= RES_PARAMS.params.get("encCd") %></td></tr>
        <tr><td class="left">mchtTrdNo[상점주문번호]</td><td class="right"><%= RES_PARAMS.params.get("mchtTrdNo") %></td></tr>
        <tr><td class="left">trdNo[헥토파이낸셜 거래번호]</td><td class="right"><%= RES_PARAMS.params.get("trdNo") %></td></tr>
        <tr><td class="left">trdDt[거래일자]</td><td class="right"><%= RES_PARAMS.params.get("trdDt") %></td></tr>
        <tr><td class="left">trdTm[거래시간]</td><td class="right"><%= RES_PARAMS.params.get("trdTm") %></td></tr>
        <tr><td class="left">outStatCd[거래상태코드]</td><td class="right"><%= RES_PARAMS.params.get("outStatCd") %></td></tr>
        <tr><td class="left">outRsltCd[거래결과코드]</td><td class="right"><%= RES_PARAMS.params.get("outRsltCd") %></td></tr>
        <tr><td class="left">outRsltMsg[결과메세지]</td><td class="right"><%= RES_PARAMS.params.get("outRsltMsg") %></td></tr>
    </table>
    <div class="tab">data</div>
    <table>
        <tr><td class="left">pktHash[해쉬값]</td><td class="right"><%= RES_PARAMS.data.get("pktHash") %></td></tr>
        <tr style="background-color:yellow;"><td class="left">billKey[삭제된 빌키]</td><td class="right"><%= RES_PARAMS.data.get("billKey") %></td></tr>
        <tr>
            <td colspan="2" style="text-align: center;"><input class="button" type="button" name="button" value="돌아가기" onclick="location.href='delKey_form.asp'"></td>
        </tr>
    </table>
</div>
</body>
</html>
<%
set REQ_PARAMS = nothing
set RES_PARAMS = nothing
set resData = nothing
%>

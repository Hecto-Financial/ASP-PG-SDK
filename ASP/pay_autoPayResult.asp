<% @CODEPAGE="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset = "UTF-8" %>
<!--#include virtual="/npg/inc/KISA_SHA256.asp"-->
<!--#include virtual="/npg/inc/settleUtils.asp"-->
<!--#include virtual="/npg/inc/config.asp"-->
<%

Dim apiHost : apiHost = CANCEL_SERVER '자동연장결제 타겟 서버



'============================================================================================================================================
'   요청 파라미터
'============================================================================================================================================
Dim REQ_PARAMS : set REQ_PARAMS = JSON.parse("{ ""params"":{}, ""data"":{} }")
REQ_PARAMS.params.set "mchtId",     "" & request.Form("mchtId")     '상점아이디
REQ_PARAMS.params.set "ver",        "" & request.Form("ver")        '버전
REQ_PARAMS.params.set "method",     "" & request.Form("method")     '결제수단
REQ_PARAMS.params.set "bizType",    "" & request.Form("bizType")    '업무구분
REQ_PARAMS.params.set "encCd",      "" & request.Form("encCd")      '암호화구분
REQ_PARAMS.params.set "mchtTrdNo",  "" & request.Form("mchtTrdNo")  '상점주문번호
REQ_PARAMS.params.set "trdDt",      "" & request.Form("trdDt")      '요청일자
REQ_PARAMS.params.set "trdTm",      "" & request.Form("trdTm")      '요청시간
REQ_PARAMS.params.set "mobileYn",   "" & request.Form("mobileYn")   '모바일여부
REQ_PARAMS.params.set "osType",     "" & request.Form("osType")     '운영체제 구분

REQ_PARAMS.data.set "telCo",    "" & request.Form("telCo")      '통신사
REQ_PARAMS.data.set "email",    "" & request.Form("email")      '상점고객이메일
REQ_PARAMS.data.set "mUserId",  "" & request.Form("mUserId")    '상점고객아이디
REQ_PARAMS.data.set "crcCd",    "" & request.Form("crcCd")      '통화구분
REQ_PARAMS.data.set "trdAmt",   "" & request.Form("trdAmt")     '거래금액
REQ_PARAMS.data.set "prdtNm",   "" & request.Form("prdtNm")     '상품명
REQ_PARAMS.data.set "sellerNm", "" & request.Form("sellerNm")   '판매자명
REQ_PARAMS.data.set "ordNm",    "" & request.Form("ordNm")      '주문자명
REQ_PARAMS.data.set "billKey",  "" & request.Form("billKey")    '자동결제키


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
RES_PARAMS.params.set "trdNo", ""       '세틀뱅크거래번호
RES_PARAMS.params.set "trdDt", ""       '요청일자
RES_PARAMS.params.set "trdTm", ""       '요청시간
RES_PARAMS.params.set "outStatCd", ""   '결과코드
RES_PARAMS.params.set "outRsltCd", ""   '거절코드
RES_PARAMS.params.set "outRsltMsg", ""  '결과메세지

RES_PARAMS.data.set "pktHash", ""       '해쉬값
RES_PARAMS.data.set "telCo", ""         '통신사
RES_PARAMS.data.set "trdAmt", ""        '거래금액
RES_PARAMS.data.set "billKey", ""       '자동결제키



'AES256 암호화 필요한 요청파라미터
Dim ENCRYPT_PARAMS : ENCRYPT_PARAMS = array("telCo", "trdAmt")

'AES256 복호화 필요한 응답파라미터
Dim DECRYPT_PARAMS : DECRYPT_PARAMS = array("telCo", "trdAmt")


'============================================================================================================================================
'   SHA256 해쉬 처리
'조합 필드 : 요청일자 + 요청시간 + 상점아이디 + 상점주문번호 + 거래금액(평문) + 라이센스키
'============================================================================================================================================
Dim hashPlain : hashPlain = ""
Dim hashCipher : hashCipher = ""
hashPlain = REQ_PARAMS.params.get("trdDt") &_
            REQ_PARAMS.params.get("trdTm") &_
            REQ_PARAMS.params.get("mchtId") &_
            REQ_PARAMS.params.get("mchtTrdNo") &_
            REQ_PARAMS.data.get("trdAmt") &_
            LICENSE_KEY
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
'   AES256 암호화 처리
'============================================================================================================================================
On error resume next
Dim objCrypto : Set objCrypto = Server.CreateObject("SBCryptoUtil.CryptoUtil.1") '암복호화 객체 생성
Dim aesPlain : aesPlain = ""
Dim aesCipher : aesCipher = ""
Dim i
for each i in ENCRYPT_PARAMS
    aesPlain = REQ_PARAMS.data.get(i)
    if "" <> aesPlain then
        aesCipher = objCrypto.EncryptBase64(AES256_KEY, aesPlain) 'AES256 암호화
        REQ_PARAMS.data.set i, aesCipher '암호화 결과 값 세팅
        call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][AES256 Encrypt] " & i & "[" & aesPlain & "] ---> [" & aesCipher & "]")
    end if
next

If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][AES256 Encrypt] Fail! " & err.number & " : " & err.description  )
End If
On Error GoTo 0



'============================================================================================================================================
'   API URL 설정
'============================================================================================================================================
Dim requestUrl : requestUrl = apiHost & "/spay/APIService.do" '휴대폰 자동연장 결제 URL
call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SEND POST] Request url :  " & requestUrl  )



'============================================================================================================================================
'   API호출(가맹점->세틀) 및 응답 처리
'============================================================================================================================================
On error resume next
call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SEND POST] Request :  " & JSON.stringify(REQ_PARAMS)  )
Dim resData : set resData = sendPost(requestUrl, JSON.stringify(REQ_PARAMS) , CONN_TIMEOUT, READ_TIMEOUT)
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
    RES_PARAMS.params.set "trdNo", resData.params.get("trdNo")              '세틀뱅크거래번호
    RES_PARAMS.params.set "trdDt", resData.params.get("trdDt")              '요청일자
    RES_PARAMS.params.set "trdTm", resData.params.get("trdTm")              '요청시간
    RES_PARAMS.params.set "outStatCd", resData.params.get("outStatCd")      '결과코드
    RES_PARAMS.params.set "outRsltCd", resData.params.get("outRsltCd")      '거절코드
    RES_PARAMS.params.set "outRsltMsg", resData.params.get("outRsltMsg")    '결과메세지
end if
if resData.get("data") <> "" then
    RES_PARAMS.data.set "pktHash", resData.data.get("pktHash")              '해쉬값
    RES_PARAMS.data.set "telCo", resData.data.get("telCo")                  '통신사
    RES_PARAMS.data.set "trdAmt", resData.data.get("trdAmt")                '거래금액
    RES_PARAMS.data.set "billKey", resData.data.get("billKey")              '자동결제키
end if



'============================================================================================================================================
'   AES256 복호화 처리
'============================================================================================================================================
On error resume next
aesPlain = ""
aesCipher = ""
for each i in DECRYPT_PARAMS
    aesCipher = Trim(RES_PARAMS.data.get(i))
    if "" <> aesCipher then
        aesPlain = objCrypto.DecryptBase64(AES256_KEY, aesCipher) 'AES256 복호화
        RES_PARAMS.data.set i, aesPlain '복호화 결과 값 세팅
        call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][AES256 Decrypt] " & i & "[" & aesCipher & "] ---> [" & aesPlain & "]")
    end if
next

If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][AES256 Decrypt] fail! " & err.number & " : " & err.description  )
End If
On Error GoTo 0


%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="ie=edge">
<title>결제 요청 결과</title>
<style type="text/css">
    body            {font-family:굴림; font-size:10pt; color:#000000; text-decoration:none;}
    font            {font-family:굴림; font-size:10pt; color:#000000; text-decoration:none;}
    td              {font-family:굴림; font-size:10pt; color:#000000; text-decoration:none; padding:3px; border:1px solid #e1e1e1;}
    .left           {padding-left:5px; width:210px;}
    .right          {padding-left:5px;}
    .wrapper        {width:700px;border:1px solid #e1e1e1;}
    .tab            {background-color:#f1f1f1;padding:10px 20px;border:1px solid #e1e1e1; font-weight: bold; font-size:1.1em;}
    table           {width:100%; border-collapse:collapse;}
    .button         {padding:5px 20px; border-radius:20px; border:1px solid #ccc; width:70%; margin:5px 0px; transition:0.3s; cursor:pointer;}
    .button:hover   {background-color:#aaaaaa;}
</style>
</head>
<body>
<h2>결제 요청 결과</h2>
<div class="wrapper">
    <div class="tab">응답 파라미터</div>
    <table>
        <tr>
            <td class="left">mchtId[상점아이디]</td>
            <td class="right"><%= RES_PARAMS.params.get("mchtId") %></td>
        </tr>
        <tr>
            <td class="left">ver[버전]</td>
            <td class="right"><%= RES_PARAMS.params.get("ver") %></td>
        </tr>
        <tr>
            <td class="left">method[결제수단]</td>
            <td class="right"><%= RES_PARAMS.params.get("method") %></td>
        </tr>
        <tr>
            <td class="left">bizType[업무구분]</td>
            <td class="right"><%= RES_PARAMS.params.get("bizType") %></td>
        </tr>
        <tr>
            <td class="left">encCd[암호화구분]</td>
            <td class="right"><%= RES_PARAMS.params.get("encCd") %></td>
        </tr>
        <tr>
            <td class="left">mchtTrdNo[상점주문번호]</td>
            <td class="right"><%= RES_PARAMS.params.get("mchtTrdNo") %></td>
        </tr>
        <tr>
            <td class="left">trdNo[세틀뱅크 거래번호]</td>
            <td class="right"><%= RES_PARAMS.params.get("trdNo") %></td>
        </tr>
        <tr>
            <td class="left">trdDt[취소요청일자]</td>
            <td class="right"><%= RES_PARAMS.params.get("trdDt") %></td>
        </tr>
        <tr>
            <td class="left">trdTm[취소요청시간]</td>
            <td class="right"><%= RES_PARAMS.params.get("trdTm") %></td>
        </tr>
        <tr>
            <td class="left">outStatCd[거래상태코드]</td>
            <td class="right"><%= RES_PARAMS.params.get("outStatCd") %></td>
        </tr>
        <tr>
            <td class="left">outRsltCd[거래결과코드]</td>
            <td class="right"><%= RES_PARAMS.params.get("outRsltCd") %></td>
        </tr>
        <tr>
            <td class="left">outRsltMsg[결과메세지]</td>
            <td class="right"><%= RES_PARAMS.params.get("outRsltMsg") %></td>
        </tr>
        <tr>
            <td class="left">pktHash[해쉬값]</td>
            <td class="right"><%= RES_PARAMS.data.get("pktHash") %></td>
        </tr>
        <tr>
            <td class="left">telCo[통신사]</td>
            <td class="right"><%= RES_PARAMS.data.get("telCo") %></td>
        </tr>
        <tr>
            <td class="left">trdAmt[거래금액]</td>
            <td class="right"><%= RES_PARAMS.data.get("trdAmt") %></td>
        </tr>
        <tr>
            <td class="left">billKey[자동결제키]</td>
            <td class="right"><%= RES_PARAMS.data.get("billKey") %></td>
        </tr>
        <tr>
            <td colspan="2" style="text-align: center;"><input class="button" type="button" name="button" value="돌아가기" onclick="location.href='pay_form.asp'"></td>
        </tr>
    </table>
</div>
</body>
</html>
<%
set REQ_PARAMS = nothing
set RES_PARAMS = nothing
set resData = nothing
set objCrypto = nothing
%>

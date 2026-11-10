<% @CODEPAGE="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset = "UTF-8" %>
<!--#include virtual="/npg/inc/KISA_SHA256.asp"-->
<!--#include virtual="/npg/inc/settleUtils.asp"-->
<!--#include virtual="/npg/inc/config.asp"-->
<%
'============================================================================================================================================
'   신용카드 빌키 결제 API 결과 처리 (2회차 결제)
'   - 호출 경로: /spay/APICardActionPay.do (게이트웨이, JSON, UI 없음 / 서버 간 통신)
'   - 빌키 결제 전용 샘플입니다. 카드번호 직접 입력(수기결제)은 지원하지 않습니다.
'============================================================================================================================================

'============================================================================================================================================
'   요청 파라미터
'============================================================================================================================================
Dim REQ_PARAMS : set REQ_PARAMS = JSON.parse("{ ""params"":{}, ""data"":{} }")
REQ_PARAMS.params.set "mchtId",     "" & request.Form("mchtId")     '상점아이디
REQ_PARAMS.params.set "ver",        "" & request.Form("ver")        '버전
REQ_PARAMS.params.set "method",     "" & request.Form("method")     '결제수단(CA)
REQ_PARAMS.params.set "bizType",    "" & request.Form("bizType")    '업무구분(B0)
REQ_PARAMS.params.set "encCd",      "" & request.Form("encCd")      '암호화구분
REQ_PARAMS.params.set "mchtTrdNo",  "" & request.Form("mchtTrdNo")  '상점주문번호
REQ_PARAMS.params.set "trdDt",      "" & request.Form("trdDt")      '요청일자
REQ_PARAMS.params.set "trdTm",      "" & request.Form("trdTm")      '요청시간
REQ_PARAMS.params.set "mobileYn",   "" & request.Form("mobileYn")   '모바일여부
REQ_PARAMS.params.set "osType",     "" & request.Form("osType")     '운영체제 구분

'요청 파라미터(바디) - 빌키 결제 전용(카드번호 등 수기 입력 필드 없음)
REQ_PARAMS.data.set "pktHash",      ""                                  '해쉬값
REQ_PARAMS.data.set "pmtprdNm",     "" & request.Form("pmtprdNm")       '상품명
REQ_PARAMS.data.set "mchtCustNm",   "" & request.Form("mchtCustNm")     '상점고객명
REQ_PARAMS.data.set "mchtCustId",   "" & request.Form("mchtCustId")     '상점고객아이디
REQ_PARAMS.data.set "email",        "" & request.Form("email")          '이메일
REQ_PARAMS.data.set "billKey",      "" & request.Form("billKey")        '빌키
REQ_PARAMS.data.set "instmtMon",    "" & request.Form("instmtMon")      '할부개월수
REQ_PARAMS.data.set "crcCd",        "" & request.Form("crcCd")          '통화구분
REQ_PARAMS.data.set "taxTypeCd",    "" & request.Form("taxTypeCd")      '세금유형
REQ_PARAMS.data.set "trdAmt",       "" & request.Form("trdAmt")         '거래금액
REQ_PARAMS.data.set "taxAmt",       "" & request.Form("taxAmt")         '과세금액
REQ_PARAMS.data.set "vatAmt",       "" & request.Form("vatAmt")         '부가세금액
REQ_PARAMS.data.set "taxFreeAmt",   "" & request.Form("taxFreeAmt")     '비과세금액
REQ_PARAMS.data.set "svcAmt",       "" & request.Form("svcAmt")         '봉사료
REQ_PARAMS.data.set "notiUrl",      "" & request.Form("notiUrl")        '결과처리URL
REQ_PARAMS.data.set "mchtParam",    "" & request.Form("mchtParam")      '기타주문정보


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
RES_PARAMS.params.set "trdDt", ""       '요청일자
RES_PARAMS.params.set "trdTm", ""       '요청시간
RES_PARAMS.params.set "outStatCd", ""   '결과코드
RES_PARAMS.params.set "outRsltCd", ""   '거절코드
RES_PARAMS.params.set "outRsltMsg", ""  '결과메세지

RES_PARAMS.data.set "pktHash", ""       '해쉬값
RES_PARAMS.data.set "trdAmt", ""        '거래금액
RES_PARAMS.data.set "billKey", ""       '빌키
RES_PARAMS.data.set "cardNo", ""        '마스킹된 카드번호(사전 협의 가맹점 한정)
RES_PARAMS.data.set "vldDtMon", ""      '유효기간(월)
RES_PARAMS.data.set "vldDtYear", ""     '유효기간(년)
RES_PARAMS.data.set "issrId", ""        '발급사아이디
RES_PARAMS.data.set "cardNm", ""        '카드사명
RES_PARAMS.data.set "cardKind", ""      '카드종류명
RES_PARAMS.data.set "ninstmtTypeCd", "" '무이자할부타입
RES_PARAMS.data.set "instmtMon", ""     '할부개월수
RES_PARAMS.data.set "apprNo", ""        '승인번호


'AES256 암호화 필요한 요청파라미터(금액 항목만)
Dim ENCRYPT_PARAMS : ENCRYPT_PARAMS = array("trdAmt", "taxAmt", "vatAmt", "taxFreeAmt", "svcAmt")

'AES256 복호화 필요한 응답파라미터
Dim DECRYPT_PARAMS : DECRYPT_PARAMS = array("trdAmt")


'============================================================================================================================================
'   SHA256 해쉬 처리
'조합 필드 : 요청일자 + 요청시간 + 상점아이디 + 상점주문번호 + 거래금액(평문) + 라이센스키
'============================================================================================================================================
Dim hashPlain : hashPlain = REQ_PARAMS.params.get("trdDt") & REQ_PARAMS.params.get("trdTm") & REQ_PARAMS.params.get("mchtId") & REQ_PARAMS.params.get("mchtTrdNo") & REQ_PARAMS.data.get("trdAmt") & LICENSE_KEY
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
'   타겟 URL 설정
'  타겟 서버 : (tb)gw.settlebank.co.kr
'  신용카드 빌키 결제 : ~/spay/APICardActionPay.do
'============================================================================================================================================
Dim requestUrl : requestUrl = CANCEL_SERVER & "/spay/APICardActionPay.do"
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
    RES_PARAMS.params.set "trdDt", resData.params.get("trdDt")              '요청일자
    RES_PARAMS.params.set "trdTm", resData.params.get("trdTm")              '요청시간
    RES_PARAMS.params.set "outStatCd", resData.params.get("outStatCd")      '결과코드
    RES_PARAMS.params.set "outRsltCd", resData.params.get("outRsltCd")      '거절코드
    RES_PARAMS.params.set "outRsltMsg", resData.params.get("outRsltMsg")    '결과메세지
end if
if resData.get("data") <> "" then
    RES_PARAMS.data.set "pktHash", resData.data.get("pktHash")              '해쉬값
    RES_PARAMS.data.set "trdAmt", resData.data.get("trdAmt")                '거래금액
    RES_PARAMS.data.set "billKey", resData.data.get("billKey")             '빌키
    RES_PARAMS.data.set "cardNo", resData.data.get("cardNo")               '마스킹된 카드번호
    RES_PARAMS.data.set "vldDtMon", resData.data.get("vldDtMon")           '유효기간(월)
    RES_PARAMS.data.set "vldDtYear", resData.data.get("vldDtYear")         '유효기간(년)
    RES_PARAMS.data.set "issrId", resData.data.get("issrId")               '발급사아이디
    RES_PARAMS.data.set "cardNm", resData.data.get("cardNm")               '카드사명
    RES_PARAMS.data.set "cardKind", resData.data.get("cardKind")           '카드종류명
    RES_PARAMS.data.set "ninstmtTypeCd", resData.data.get("ninstmtTypeCd") '무이자할부타입
    RES_PARAMS.data.set "instmtMon", resData.data.get("instmtMon")         '할부개월수
    RES_PARAMS.data.set "apprNo", resData.data.get("apprNo")               '승인번호
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
<title>빌키 결제 응답 결과</title>
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
<h2>빌키 결제 응답 결과</h2>
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
        <tr><td class="left">trdDt[요청일자]</td><td class="right"><%= RES_PARAMS.params.get("trdDt") %></td></tr>
        <tr><td class="left">trdTm[요청시간]</td><td class="right"><%= RES_PARAMS.params.get("trdTm") %></td></tr>
        <tr><td class="left">outStatCd[거래상태코드]</td><td class="right"><%= RES_PARAMS.params.get("outStatCd") %></td></tr>
        <tr><td class="left">outRsltCd[거래결과코드]</td><td class="right"><%= RES_PARAMS.params.get("outRsltCd") %></td></tr>
        <tr><td class="left">outRsltMsg[결과메세지]</td><td class="right"><%= RES_PARAMS.params.get("outRsltMsg") %></td></tr>
    </table>
    <div class="tab">data</div>
    <table>
        <tr><td class="left">pktHash[해쉬값]</td><td class="right"><%= RES_PARAMS.data.get("pktHash") %></td></tr>
        <tr><td class="left">trdAmt[거래금액]</td><td class="right"><%= RES_PARAMS.data.get("trdAmt") %></td></tr>
        <tr><td class="left">cardNo[카드번호(마스킹)]</td><td class="right"><%= RES_PARAMS.data.get("cardNo") %></td></tr>
        <tr><td class="left">vldDtYear[유효기간(년)]</td><td class="right"><%= RES_PARAMS.data.get("vldDtYear") %></td></tr>
        <tr><td class="left">vldDtMon[유효기간(월)]</td><td class="right"><%= RES_PARAMS.data.get("vldDtMon") %></td></tr>
        <tr><td class="left">issrId[발급사아이디]</td><td class="right"><%= RES_PARAMS.data.get("issrId") %></td></tr>
        <tr><td class="left">cardNm[카드사명]</td><td class="right"><%= RES_PARAMS.data.get("cardNm") %></td></tr>
        <tr><td class="left">cardKind[카드종류명]</td><td class="right"><%= RES_PARAMS.data.get("cardKind") %></td></tr>
        <tr><td class="left">ninstmtTypeCd[무이자할부타입]</td><td class="right"><%= RES_PARAMS.data.get("ninstmtTypeCd") %></td></tr>
        <tr><td class="left">instmtMon[할부개월수]</td><td class="right"><%= RES_PARAMS.data.get("instmtMon") %></td></tr>
        <tr><td class="left">apprNo[승인번호]</td><td class="right"><%= RES_PARAMS.data.get("apprNo") %></td></tr>
        <tr style="background-color:yellow;"><td class="left">billKey[빌키]</td><td class="right"><%= RES_PARAMS.data.get("billKey") %></td></tr>
        <tr>
            <td colspan="2" style="text-align: center;"><input class="button" type="button" name="button" value="돌아가기" onclick="location.href='billkey_pay_form.asp'"></td>
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

<% @CODEPAGE="65001" language="VBScript" %>
<%
Response.CharSet="utf-8"
Session.codepage="65001"
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"
%>
<!--#include virtual="/npg/inc/config.asp"-->
<!--#include virtual="/npg/inc/KISA_SHA256.asp"-->
<!--#include virtual="/npg/inc/settleUtils.asp"-->
<%

'응답 파라미터 세팅
Dim RES_PARAMS : set RES_PARAMS = Server.CreateObject("Scripting.Dictionary")
RES_PARAMS.Add "mchtId",        "" & request.form("mchtId")         '상점아이디
RES_PARAMS.Add "outStatCd",     "" & request.form("outStatCd")      '결과코드
RES_PARAMS.Add "outRsltCd",     "" & request.form("outRsltCd")      '거절코드
RES_PARAMS.Add "outRsltMsg",    "" & request.form("outRsltMsg")     '결과메세지
RES_PARAMS.Add "method",        "" & request.form("method")         '결제수단
RES_PARAMS.Add "mchtTrdNo",     "" & request.form("mchtTrdNo")      '상점주문번호
RES_PARAMS.Add "mchtCustId",    "" & request.form("mchtCustId")     '상점고객아이디
RES_PARAMS.Add "trdNo",         "" & request.form("trdNo")          '헥토파이낸셜거래번호
RES_PARAMS.Add "trdAmt",        "" & request.form("trdAmt")         '거래금액
RES_PARAMS.Add "mchtParam",     "" & request.form("mchtParam")      '상점예약필드
RES_PARAMS.Add "authDt",        "" & request.form("authDt")         '승인일시
RES_PARAMS.Add "authNo",        "" & request.form("authNo")         '승인번호
RES_PARAMS.Add "reqIssueDt",    "" & request.form("reqIssueDt")     '채번요청일시
RES_PARAMS.Add "intMon",        "" & request.form("intMon")         '할부개월수
RES_PARAMS.Add "fnNm",          "" & request.form("fnNm")           '카드사명
RES_PARAMS.Add "fnCd",          "" & request.form("fnCd")           '카드사코드
RES_PARAMS.Add "pointTrdNo",    "" & request.form("pointTrdNo")     '포인트거래번호
RES_PARAMS.Add "pointTrdAmt",   "" & request.form("pointTrdAmt")    '포인트거래금액
RES_PARAMS.Add "cardTrdAmt",    "" & request.form("cardTrdAmt")     '신용카드결제금액
RES_PARAMS.Add "vtlAcntNo",     "" & request.form("vtlAcntNo")      '가상계좌번호
RES_PARAMS.Add "expireDt",      "" & request.form("expireDt")       '입금기한
RES_PARAMS.Add "cphoneNo",      "" & request.form("cphoneNo")       '휴대폰번호
RES_PARAMS.Add "billKey",       "" & request.form("billKey")        '자동결제키
RES_PARAMS.Add "csrcAmt",       "" & request.form("csrcAmt")        '현금영수증 발급 금액(네이버페이)

'AES256 복호화 필요 파라미터
Dim DECRYPT_PARAMS : DECRYPT_PARAMS = array("mchtCustId","trdAmt", "pointTrdAmt", "cardTrdAmt", "vtlAcntNo", "cphoneNo", "csrcAmt")


'============================================================================================================================================
'   AES256 복호화 처리(Base64 decoding -> AES-256-ECB decrypt )
'============================================================================================================================================
On error resume next
Dim objCrypto : Set objCrypto = Server.CreateObject("SBCryptoUtil.CryptoUtil.1") '암복호화 객체 생성

Dim aesPlain : aesPlain = ""
Dim aesCipher : aesCipher = ""
for each i in DECRYPT_PARAMS
    aesCipher = Trim(RES_PARAMS.Item(i))
    if "" <> aesCipher then
        aesPlain = objCrypto.DecryptBase64(AES256_KEY, aesCipher) 'AES256 복호화
        RES_PARAMS.Item(i) = aesPlain '복호화 결과 값 세팅
        call log_message(LOG_FILE, "[" & RES_PARAMS.Item("mchtTrdNo") & "][AES256 Decrypt] " & i & "[" & aesCipher & "] ---> [" & aesPlain & "]")
    end if
next

If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & RES_PARAMS.Item("mchtTrdNo") & "][AES256 Decrypt] fail! " & err.number & " : " & err.description  )
End If
On Error GoTo 0

'응답 파라미터 로깅
Dim logStr : logStr = "[" & RES_PARAMS.Item("mchtTrdNo") & "][Response Data] "
Dim key, val, idx
key = RES_PARAMS.Keys
val = RES_PARAMS.Items
For idx = 0 To RES_PARAMS.Count -1
    logStr = logStr & key(idx) & "(" & val(idx) & ") "
Next
call log_message(LOG_FILE, logStr)

%>
<html>
<head><title>헥토파이낸셜 결제 결과 페이지</title>
<meta name="viewport" content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="ie=edge">
<style type="text/css">
    body            {font-family:굴림; font-size:10pt; color:#000000; text-decoration:none;}
    font            {font-family:굴림; font-size:10pt; color:#000000; text-decoration:none;}
    td              {font-family:굴림; font-size:10pt; color:#000000; text-decoration:none; padding:3px; border:1px solid #e1e1e1;}
    .left           {padding-left:5px; width:100px;}
    .right          {padding-left:5px;}
    .wrapper        {max-width:700px;border:1px solid #e1e1e1;}
    .tab            {background-color:#f1f1f1;padding:10px 20px;border:1px solid #e1e1e1; font-weight: bold; font-size:1.1em;}
    table           {width:100%; border-collapse:collapse;}
    .button         {padding:5px 20px; border-radius:20px; border:1px solid #ccc; width:70%; margin:5px 0px; transition:0.3s; cursor:pointer;}
    .button:hover   {background-color:#aaaaaa;}
</style>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
<script>
//결제 결과 세팅
var _PAY_RESULT = {
    mchtId :        "<%= RES_PARAMS.Item("mchtId") %>",
    outStatCd :     "<%= RES_PARAMS.Item("outStatCd") %>",
    outRsltCd :     "<%= RES_PARAMS.Item("outRsltCd") %>",
    outRsltMsg :    "<%= RES_PARAMS.Item("outRsltMsg") %>",
    method :        "<%= RES_PARAMS.Item("method") %>",
    mchtTrdNo :     "<%= RES_PARAMS.Item("mchtTrdNo") %>",
    mchtCustId :    "<%= RES_PARAMS.Item("mchtCustId") %>",
    trdNo :         "<%= RES_PARAMS.Item("trdNo") %>",
    trdAmt :        "<%= RES_PARAMS.Item("trdAmt") %>",
    mchtParam :     "<%= RES_PARAMS.Item("mchtParam") %>",
    authDt :        "<%= RES_PARAMS.Item("authDt") %>",
    authNo :        "<%= RES_PARAMS.Item("authNo") %>",
    reqIssueDt :    "<%= RES_PARAMS.Item("reqIssueDt") %>",
    intMon :        "<%= RES_PARAMS.Item("intMon") %>",
    fnNm :          "<%= RES_PARAMS.Item("fnNm") %>",
    fnCd :          "<%= RES_PARAMS.Item("fnCd") %>",
    pointTrdNo :    "<%= RES_PARAMS.Item("pointTrdNo") %>",
    pointTrdAmt :   "<%= RES_PARAMS.Item("pointTrdAmt") %>",
    cardTrdAmt :    "<%= RES_PARAMS.Item("cardTrdAmt") %>",
    vtlAcntNo :     "<%= RES_PARAMS.Item("vtlAcntNo") %>",
    expireDt :      "<%= RES_PARAMS.Item("expireDt") %>",
    cphoneNo :      "<%= RES_PARAMS.Item("cphoneNo") %>",
    billKey :       "<%= RES_PARAMS.Item("billKey") %>",
    csrcAmt :       "<%= RES_PARAMS.Item("csrcAmt") %>"
};

//main으로 결과 전달
function sendResult()
{
    if(top.opener){
        //팝업창
        top.opener.rstparamSet(_PAY_RESULT);
        top.opener.goResult();
        self.close();
    }
    else{//iframe
        parent.postMessage(JSON.stringify({action:"HECTO_IFRAME_CLOSE", params: _PAY_RESULT}), "*");
    }
}
</script>
</head>
<body>
<h2>승인 요청 결과</h2>
<div class="wrapper">
    <div class="tab">응답 파라미터</div>
    <table>
        <tr>
            <td class="left">mchtId</td>
            <td class="right"><%= RES_PARAMS.Item("mchtId") %></td>
        </tr>
        <tr>
            <td class="left">outStatCd</td>
            <td class="right"><%= RES_PARAMS.Item("outStatCd") %></td>
        </tr>
        <tr>
            <td class="left">outRsltCd</td>
            <td class="right"><%= RES_PARAMS.Item("outRsltCd") %></td>
        </tr>
        <tr>
            <td class="left">outRsltMsg</td>
            <td class="right"><%= RES_PARAMS.Item("outRsltMsg") %></td>
        </tr>
        <tr>
            <td class="left">method</td>
            <td class="right"><%= RES_PARAMS.Item("method") %></td>
        </tr>
        <tr>
            <td class="left">mchtTrdNo</td>
            <td class="right"><%= RES_PARAMS.Item("mchtTrdNo") %></td>
        </tr>
        <tr>
            <td class="left">mchtCustId</td>
            <td class="right"><%= RES_PARAMS.Item("mchtCustId") %></td>
        </tr>
        <tr>
            <td class="left">trdNo</td>
            <td class="right"><%= RES_PARAMS.Item("trdNo") %></td>
        </tr>
        <tr>
            <td class="left">trdAmt</td>
            <td class="right"><%= RES_PARAMS.Item("trdAmt") %></td>
        </tr>
        <tr>
            <td class="left">mchtParam</td>
            <td class="right"><%= RES_PARAMS.Item("mchtParam") %></td>
        </tr>
        <tr>
            <td class="left">authDt</td>
            <td class="right"><%= RES_PARAMS.Item("authDt") %></td>
        </tr>
        <tr>
            <td class="left">authNo</td>
            <td class="right"><%= RES_PARAMS.Item("authNo") %></td>
        </tr>
        <tr>
            <td class="left">reqIssueDt</td>
            <td class="right"><%= RES_PARAMS.Item("reqIssueDt") %></td>
        </tr>
        <tr>
            <td class="left">intMon</td>
            <td class="right"><%= RES_PARAMS.Item("intMon") %></td>
        </tr>
        <tr>
            <td class="left">fnNm</td>
            <td class="right"><%= RES_PARAMS.Item("fnNm") %></td>
        </tr>
        <tr>
            <td class="left">fnCd</td>
            <td class="right"><%= RES_PARAMS.Item("fnCd") %></td>
        </tr>
        <tr>
            <td class="left">pointTrdNo</td>
            <td class="right"><%= RES_PARAMS.Item("pointTrdNo") %></td>
        </tr>
        <tr>
            <td class="left">pointTrdAmt</td>
            <td class="right"><%= RES_PARAMS.Item("pointTrdAmt") %></td>
        </tr>
        <tr>
            <td class="left">cardTrdAmt</td>
            <td class="right"><%= RES_PARAMS.Item("cardTrdAmt") %></td>
        </tr>
        <tr>
            <td class="left">vtlAcntNo</td>
            <td class="right"><%= RES_PARAMS.Item("vtlAcntNo") %></td>
        </tr>
        <tr>
            <td class="left">expireDt</td>
            <td class="right"><%= RES_PARAMS.Item("expireDt") %></td>
        </tr>
        <tr>
            <td class="left">cphoneNo</td>
            <td class="right"><%= RES_PARAMS.Item("cphoneNo") %></td>
        </tr>
        <tr>
            <td class="left">billKey</td>
            <td class="right"><%= RES_PARAMS.Item("billKey") %></td>
        </tr>
        <tr>
            <td class="left">csrcAmt</td>
            <td class="right"><%= RES_PARAMS.Item("csrcAmt") %></td>
        </tr>
        

        <tr>
            <td colspan="2" style="text-align: center;">
                <input class="button" type="button" value="확인" onclick="sendResult()" /> 
            </td>
        </tr>
    </table>
</div>
</body>
</html>
<%
set objCrypto = nothing
set RES_PARAMS = nothing
set DECRYPT_PARAMS = nothing
%>
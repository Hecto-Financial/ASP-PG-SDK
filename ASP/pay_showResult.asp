<% @CODEPAGE="65001" language="VBScript" %>
<%
Response.CharSet="utf-8"
Session.codepage="65001"
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"
%>
<!--#include virtual="/npg/inc/config.asp"-->
<!--#include virtual="/npg/inc/settleUtils.asp"-->
<%
'넘어온 응답 파라미터 받기
Dim mchtId : mchtId                     = "" & request.Form("respMchtId")           '상점아이디
Dim outStatCd : outStatCd               = "" & request.Form("respOutStatCd")        '결과코드
Dim outRsltCd : outRsltCd               = "" & request.Form("respOutRsltCd")        '거절코드
Dim outRsltMsg : outRsltMsg             = "" & request.Form("respOutRsltMsg")       '결과메세지
Dim method : method                     = "" & request.Form("respMethod")           '결제수단
Dim mchtTrdNo : mchtTrdNo               = "" & request.Form("respMchtTrdNo")        '상점주문번호
Dim mchtCustId : mchtCustId             = "" & request.Form("respMchtCustId")       '상점고객아이디
Dim trdNo : trdNo                       = "" & request.Form("respTrdNo")            '헥토파이낸셜 거래번호
Dim trdAmt : trdAmt                     = "" & request.Form("respTrdAmt")           '거래금액
Dim mchtParam : mchtParam               = "" & request.Form("respMchtParam")        '상점예약필드
Dim authDt : authDt                     = "" & request.Form("respAuthDt")           '승인일시
Dim authNo : authNo                     = "" & request.Form("respAuthNo")           '승인번호
Dim reqIssueDt : reqIssueDt             = "" & request.Form("respReqIssueDt")       '채번요청일시
Dim intMon : intMon                     = "" & request.Form("respIntMon")           '할부개월수
Dim fnNm : fnNm                         = "" & request.Form("respFnNm")             '카드사명
Dim fnCd : fnCd                         = "" & request.Form("respFnCd")             '카드사코드
Dim pointTrdNo : pointTrdNo             = "" & request.Form("respPointTrdNo")       '포인트거래번호
Dim pointTrdAmt : pointTrdAmt           = "" & request.Form("respPointTrdAmt")      '포인트거래금액
Dim cardTrdAmt : cardTrdAmt             = "" & request.Form("respCardTrdAmt")       '신용카드결제금액
Dim vtlAcntNo : vtlAcntNo               = "" & request.Form("respVtlAcntNo")        '가상계좌번호
Dim expireDt : expireDt                 = "" & request.Form("respExpireDt")         '입금만료일시
Dim cphoneNo : cphoneNo                 = "" & request.Form("respCphoneNo")         '휴대폰번호
Dim billKey : billKey                   = "" & request.Form("respBillKey")          '자동결제키(휴대폰)
Dim csrcAmt : csrcAmt                   = "" & request.Form("respCsrcAmt")          '현금영수증 발급 금액(네이버페이)
%>
<html>
<head>
<meta name="viewport" content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="ie=edge">
<title>헥토파이낸셜 결제 결과 페이지</title>
<style type="text/css">
    body            {font-family:굴림; font-size:10pt; color:#000000; text-decoration:none;}
    font            {font-family:굴림; font-size:10pt; color:#000000; text-decoration:none;}
    td              {font-family:굴림; font-size:10pt; color:#000000; text-decoration:none; padding:3px; border:1px solid #e1e1e1;}
    table           {width:100%; border-collapse:collapse;}
    .left           {padding-left:5px; width:200px;}
    .right          {padding-left:5px;}
    .wrapper        {width:700px;border:1px solid #e1e1e1;}
    .tab            {background-color:#f1f1f1;padding:10px 20px;border:1px solid #e1e1e1; font-weight: bold; font-size:1.1em;}
    .button         {padding:5px 20px; border-radius:20px; border:1px solid #ccc; width:70%; margin:5px 0px; transition:0.3s; cursor:pointer;}
    .button:hover   {background-color:#aaaaaa;}
</style>
</head>
<body>
<h2>승인 요청 결과</h2>
<div class="wrapper">
    <div class="tab">응답 파라미터</div>
    <table>
        <tr>
            <td class="left">mchtId[상점아이디]</td>
            <td class="right"><%= Server.HTMLEncode(mchtId) %></td>
        </tr>
        <tr>
            <td class="left">outStatCd[거래상태]</td>
            <td class="right"><%= Server.HTMLEncode(outStatCd) %></td>
        </tr>
        <tr>
            <td class="left">outRsltCd[거절코드]</td>
            <td class="right"><%= Server.HTMLEncode(outRsltCd) %></td>
        </tr>
        <tr>
            <td class="left">outRsltMsg[메세지]</td>
            <td class="right"><%= Server.HTMLEncode(outRsltMsg) %></td>
        </tr>
        <tr>
            <td class="left">method[결제수단]</td>
            <td class="right"><%= Server.HTMLEncode(method) %></td>
        </tr>
        <tr>
            <td class="left">mchtTrdNo[상점주문번호]</td>
            <td class="right"><%= Server.HTMLEncode(mchtTrdNo) %></td>
        </tr>
        <tr>
            <td class="left">mchtCustId[상점고객아이디]</td>
            <td class="right"><%= Server.HTMLEncode(mchtCustId) %></td>
        </tr>
        <tr>
            <td class="left">trdNo[헥토파이낸셜거래번호]</td>
            <td class="right"><%= Server.HTMLEncode(trdNo) %></td>
        </tr>
        <tr>
            <td class="left">trdAmt[거래금액]</td>
            <td class="right"><%= Server.HTMLEncode(trdAmt) %></td>
        </tr>
        <tr>
            <td class="left">mchtParam[상점예약필드]</td>
            <td class="right"><%= Server.HTMLEncode(mchtParam) %></td>
        </tr>

        <tr>
            <td class="left">authDt[승인일시]</td>
            <td class="right"><%= Server.HTMLEncode(authDt) %></td>
        </tr>
        <tr>
            <td class="left">authNo[승인번호]</td>
            <td class="right"><%= Server.HTMLEncode(authNo) %></td>
        </tr>
        <tr>
            <td class="left">reqIssueDt[채번요청일시]</td>
            <td class="right"><%= Server.HTMLEncode(reqIssueDt) %></td>
        </tr>
        <tr>
            <td class="left">intMon[할부개월수]</td>
            <td class="right"><%= Server.HTMLEncode(intMon) %></td>
        </tr>
        <tr>
            <td class="left">fnNm[카드사명]</td>
            <td class="right"><%= Server.HTMLEncode(fnNm) %></td>
        </tr>
        <tr>
            <td class="left">fnCd[카드사코드]</td>
            <td class="right"><%= Server.HTMLEncode(fnCd) %></td>
        </tr>
        <tr>
            <td class="left">pointTrdNo[포인트거래번호]</td>
            <td class="right"><%= Server.HTMLEncode(pointTrdNo) %></td>
        </tr>
        <tr>
            <td class="left">pointTrdAmt[포인트거래금액]</td>
            <td class="right"><%= Server.HTMLEncode(pointTrdAmt) %></td>
        </tr>
        <tr>
            <td class="left">cardTrdAmt[신용카드결제금액]</td>
            <td class="right"><%= Server.HTMLEncode(cardTrdAmt) %></td>
        </tr>
        <tr>
            <td class="left">vtlAcntNo[가상계좌번호]</td>
            <td class="right"><%= Server.HTMLEncode(vtlAcntNo) %></td>
        </tr>
        <tr>
            <td class="left">expireDt[입금기한]</td>
            <td class="right"><%= Server.HTMLEncode(expireDt) %></td>
        </tr>
        <tr>
            <td class="left">cphoneNo[휴대폰번호]</td>
            <td class="right"><%= Server.HTMLEncode(cphoneNo) %></td>
        </tr>
        <tr>
            <td class="left">billKey[자동결제키]</td>
            <td class="right"><%= Server.HTMLEncode(billKey) %></td>
        </tr>
        <tr>
            <td class="left">csrcAmt[현금영수증 발급 금액(네이버페이)]</td>
            <td class="right"><%= Server.HTMLEncode(csrcAmt) %></td>
        </tr>
        
        <tr>
            <td colspan="2" style="text-align: center;">
                <input class="button" type="button" name="button" value="돌아가기" onclick="location.href='pay_form.asp'">
            </td>
        </tr>
    </table>
</div>
</body>
</html>
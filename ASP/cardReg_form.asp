<% @CODEPAGE="65001" language="VBScript" %>
<%
Response.CharSet="utf-8"
Session.codepage="65001"
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"
%>
<!--#include virtual="/npg/inc/config.asp"-->

<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="ie=edge">
<title>헥토파이낸셜 PG 신용카드 0원 등록(빌키발급) 샘플</title>
<style type="text/css">
    body            {font-family:굴림; font-size:10pt; color:#000000;}
    td              {font-family:굴림; font-size:10pt; padding:1px; border:1px solid #e1e1e1;}
    table           {width:100%; border-collapse:collapse;}
    .left           {padding-left:10px; width:30%;}
    .right          {width:70%;}
    .right span     {font-size:9pt; color:red; padding:0 2px;}
    input[type='text'] {width:350px; margin:5px;}
    form            {margin:0;}
    .wrapper        {width:470px; border:1px solid #e1e1e1; margin-top:20px;}
    .tab            {background-color:#f1f1f1; padding:10px 20px; border:1px solid #e1e1e1; font-weight:bold; font-size:1.1em;}
    .content        {padding:3px;}
    .payBtn         {background-color:#fff; border-radius:8px; border:1px solid #ccc; height:50px; width:99%; margin:3px 0; cursor:pointer;}
    .payBtn:hover   {background-color:#ddd;}
</style>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
<script type="text/javascript" src="<%= PAYMENT_SERVER %>/resources/js/v1/SettlePG_v1.2.js"></script>
<script type="text/javascript">
/**
 * 신용카드 0원 등록(빌키발급) 결제창
 * - 실제 결제(승인) 없이 0원 인증으로 카드를 등록하고 정기결제용 빌키(billKey)를 발급받습니다.
 * - 결제창(UI) 방식이며, 내부적으로 /card/cardAuth.do 를 호출합니다.
 *   (JS SDK: method="card" + methodSub="auth" 조합이 빌키발급창으로 라우팅됩니다.)
 * - 발급받은 billKey 로 2회차 결제는 billkey_pay_form.asp(빌키 결제 API)를 사용하세요.
 * - 거래금액(trdAmt)은 0원 인증 방식이므로 "0" 고정입니다.
 */

/** 날짜/주문번호/고객IP 세팅 */
function init(){
    var curr_date = new Date();
    var year = curr_date.getFullYear().toString();
    var month = ("0" + (curr_date.getMonth() + 1)).slice(-2).toString();
    var day = ("0" + (curr_date.getDate())).slice(-2).toString();
    var hours = ("0" + curr_date.getHours()).slice(-2).toString();
    var mins = ("0" + curr_date.getMinutes()).slice(-2).toString();
    var secs = ("0" + curr_date.getSeconds()).slice(-2).toString();
    var random4 = ("000" + Math.random() * 10000 ).slice(-4).toString();

    $('#STPG_cardRegForm [name="custIp"]').val("<%= Request.ServerVariables("REMOTE_ADDR") %>"); //고객 IP 세팅
    $('#STPG_cardRegForm [name="trdDt"]').val(year + month + day);  //요청일자 세팅
    $('#STPG_cardRegForm [name="trdTm"]').val(hours + mins + secs); //요청시간 세팅
    $('#STPG_cardRegForm [name="mchtTrdNo"]').val("CARDREG" + year + month + day + hours + mins + secs + random4); //주문번호 세팅
}

/** 카드 등록(빌키발급) 버튼 동작 */
function registerCard(){
    init();

    //용도 : SHA256 해쉬 처리 및 민감정보 AES256암호화 (pay_encryptParams.asp 재사용)
    //해시 조합 : 상점아이디 + 결제수단 + 상점주문번호 + 요청일자 + 요청시간 + 거래금액(평문,"0") + 라이센스키
    $.ajax({
        type : "POST",
        url : "pay_encryptParams.asp",
        dataType : "json",
        data : $("#STPG_cardRegForm").serialize(),
        success : function(rsp){
            //암호화 된 파라미터 세팅
            for(name in rsp.encParams) {
                $('#STPG_cardRegForm [name='+name+']').val( rsp.encParams[name] );
            };

            //가맹점 -> 헥토파이낸셜 빌키발급 결제창 호출
            SETTLE_PG.pay({
                env : "<%= PAYMENT_SERVER %>",   //결제서버 URL
                mchtId : $('#STPG_cardRegForm [name="mchtId"]').val(),
                method : "card",        //신용카드 (고정)
                methodSub : "auth",     //빌키발급창(0원 등록) (고정)
                trdDt : $('#STPG_cardRegForm [name="trdDt"]').val(),
                trdTm : $('#STPG_cardRegForm [name="trdTm"]').val(),
                mchtTrdNo : $('#STPG_cardRegForm [name="mchtTrdNo"]').val(),
                mchtName : $('#STPG_cardRegForm [name="mchtName"]').val(),
                mchtEName : $('#STPG_cardRegForm [name="mchtEName"]').val(),
                pmtPrdtNm : $('#STPG_cardRegForm [name="pmtPrdtNm"]').val(),
                trdAmt : $('#STPG_cardRegForm [name="trdAmt"]').val(),   //암호문("0" 암호화)
                mchtCustNm : $('#STPG_cardRegForm [name="mchtCustNm"]').val(),
                mchtCustId : $('#STPG_cardRegForm [name="mchtCustId"]').val(),
                notiUrl : $('#STPG_cardRegForm [name="notiUrl"]').val(),
                nextUrl : $('#STPG_cardRegForm [name="nextUrl"]').val(),
                cancUrl : $('#STPG_cardRegForm [name="cancUrl"]').val(),
                mchtParam : $('#STPG_cardRegForm [name="mchtParam"]').val(),
                custIp : $('#STPG_cardRegForm [name="custIp"]').val(),
                pktHash : rsp.hashCipher,   //SHA256 처리된 해쉬 값 세팅
                ui : {
                    type : "popup",   //popup, iframe, self, blank
                    width : "430",
                    height : "660"
                }
            }, function(rsp){
                console.log(rsp);
            });
        },
        error : function(){ alert("암호화 처리 중 에러가 발생했습니다."); }
    });
}

/** pay_receiveResult.asp(팝업)로부터 응답 값을 받아 폼에 세팅 */
function rstparamSet(rslt){
    $('#STPG_cardRegForm [name="respMchtId"]').val(rslt.mchtId);
    $('#STPG_cardRegForm [name="respOutStatCd"]').val(rslt.outStatCd);
    $('#STPG_cardRegForm [name="respOutRsltCd"]').val(rslt.outRsltCd);
    $('#STPG_cardRegForm [name="respOutRsltMsg"]').val(rslt.outRsltMsg);
    $('#STPG_cardRegForm [name="respMethod"]').val(rslt.method);
    $('#STPG_cardRegForm [name="respMchtTrdNo"]').val(rslt.mchtTrdNo);
    $('#STPG_cardRegForm [name="respMchtCustId"]').val(rslt.mchtCustId);
    $('#STPG_cardRegForm [name="respTrdNo"]').val(rslt.trdNo);
    $('#STPG_cardRegForm [name="respTrdAmt"]').val(rslt.trdAmt);
    $('#STPG_cardRegForm [name="respMchtParam"]').val(rslt.mchtParam);
    $('#STPG_cardRegForm [name="respAuthDt"]').val(rslt.authDt);
    $('#STPG_cardRegForm [name="respAuthNo"]').val(rslt.authNo);
    $('#STPG_cardRegForm [name="respReqIssueDt"]').val(rslt.reqIssueDt);
    $('#STPG_cardRegForm [name="respIntMon"]').val(rslt.intMon);
    $('#STPG_cardRegForm [name="respFnNm"]').val(rslt.fnNm);
    $('#STPG_cardRegForm [name="respFnCd"]').val(rslt.fnCd);
    $('#STPG_cardRegForm [name="respPointTrdNo"]').val(rslt.pointTrdNo);
    $('#STPG_cardRegForm [name="respPointTrdAmt"]').val(rslt.pointTrdAmt);
    $('#STPG_cardRegForm [name="respCardTrdAmt"]').val(rslt.cardTrdAmt);
    $('#STPG_cardRegForm [name="respVtlAcntNo"]').val(rslt.vtlAcntNo);
    $('#STPG_cardRegForm [name="respExpireDt"]').val(rslt.expireDt);
    $('#STPG_cardRegForm [name="respCphoneNo"]').val(rslt.cphoneNo);
    $('#STPG_cardRegForm [name="respBillKey"]').val(rslt.billKey);
    $('#STPG_cardRegForm [name="respCsrcAmt"]').val(rslt.csrcAmt);
}

/** 폼에 세팅된 응답 값을 결과 페이지로 출력 */
function goResult(){
    $('#STPG_cardRegForm').attr("action", "pay_showResult.asp");
    $('#STPG_cardRegForm').attr("method", "post");
    $('#STPG_cardRegForm').attr("target", "");
    $('#STPG_cardRegForm').submit();
}
</script>
</head>
<body>
<button onclick="location.href='pay_form.asp'">결제 샘플로</button>
<button onclick="location.href='billkey_pay_form.asp'">빌키 결제 샘플로</button>
<div class="wrapper">
    <form id="STPG_cardRegForm" name="STPG_cardRegForm">
        <!-- 요청 파라미터(필수) -->
        <input type="hidden" name="method" value="card" />
        <input type="hidden" name="trdDt" value="" />
        <input type="hidden" name="trdTm" value="" />
        <input type="hidden" name="mchtTrdNo" value="" />
        <input type="hidden" name="mchtName" value="헥토파이낸셜" />
        <input type="hidden" name="mchtEName" value="HectoFinancial" />
        <input type="hidden" name="pmtPrdtNm" value="테스트상품" />
        <input type="hidden" name="notiUrl" value="http://localhost/receiveNoti.asp" />
        <input type="hidden" name="nextUrl" value="http://localhost/pay_receiveResult.asp" />
        <input type="hidden" name="cancUrl" value="http://localhost/pay_receiveResult.asp" />
        <input type="hidden" name="custIp" value="" />
        <input type="hidden" name="mchtParam" value="상점 예약 필드" />

        <!-- 암호화 처리 후 세팅될 파라미터 -->
        <input type="hidden" name="trdAmt" />
        <input type="hidden" name="mchtCustNm" />
        <input type="hidden" name="mchtCustId" />

        <!-- 응답 파라미터(팝업 복귀 시 세팅) -->
        <input type="hidden" name="respMchtId" />
        <input type="hidden" name="respOutStatCd" />
        <input type="hidden" name="respOutRsltCd" />
        <input type="hidden" name="respOutRsltMsg" />
        <input type="hidden" name="respMethod" />
        <input type="hidden" name="respMchtTrdNo" />
        <input type="hidden" name="respMchtCustId" />
        <input type="hidden" name="respTrdNo" />
        <input type="hidden" name="respTrdAmt" />
        <input type="hidden" name="respMchtParam" />
        <input type="hidden" name="respAuthDt" />
        <input type="hidden" name="respAuthNo" />
        <input type="hidden" name="respReqIssueDt" />
        <input type="hidden" name="respIntMon" />
        <input type="hidden" name="respFnNm" />
        <input type="hidden" name="respFnCd" />
        <input type="hidden" name="respPointTrdNo" />
        <input type="hidden" name="respPointTrdAmt" />
        <input type="hidden" name="respCardTrdAmt" />
        <input type="hidden" name="respVtlAcntNo" />
        <input type="hidden" name="respExpireDt" />
        <input type="hidden" name="respCphoneNo" />
        <input type="hidden" name="respBillKey" />
        <input type="hidden" name="respCsrcAmt" />

        <div class="tab">신용카드 0원 등록(빌키발급) 결제창 호출</div>
        <table>
            <tr>
                <td class="left">상점아이디</td>
                <td class="right">
                    <input type="text" name="mchtId" value="nxca_jt_gu" />
                    <span>구인증 nxca_jt_gu / 비인증 nxca_jt_bi</span>
                </td>
            </tr>
            <tr>
                <td class="left">거래금액</td>
                <td class="right">
                    <input type="text" name="plainTrdAmt" value="0" readonly />
                    <span>0원 인증 고정</span>
                </td>
            </tr>
            <tr>
                <td class="left">고객명</td>
                <td class="right"><input type="text" name="plainMchtCustNm" value="홍길동" /></td>
            </tr>
            <tr>
                <td class="left">상점고객아이디</td>
                <td class="right"><input type="text" name="plainMchtCustId" value="HongGilDong" /></td>
            </tr>
        </table>
        <div class="content">
            <input type="button" class="payBtn" value="신용카드 등록(빌키발급)" onclick="registerCard()" />
        </div>
    </form>
</div>
</body>
</html>

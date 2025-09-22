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

'해쉬 및 aes256암호화 후 리턴 될 json
Dim rsp : set rsp = JSON.parse("{}")


'SHA256 해쉬 파라미터
Dim mchtId : mchtId         = request.form("mchtId")
Dim method : method         = request.form("method")
Dim mchtTrdNo : mchtTrdNo   = request.form("mchtTrdNo")
Dim trdDt : trdDt           = request.form("trdDt")
Dim trdTm : trdTm           = request.form("trdTm")
Dim trdAmt : trdAmt         = request.form("plainTrdAmt")


'AES256 암호화 파라미터
Dim params : set params = CreateObject("Scripting.Dictionary")
params.Add "trdAmt",            trdAmt
params.Add "mchtCustNm",        request.form("plainMchtCustNm")
params.Add "cphoneNo",          request.form("plainCphoneNo")
params.Add "email",             request.form("plainEmail")
params.Add "mchtCustId",        request.form("plainMchtCustId")
params.Add "taxAmt",            request.form("plainTaxAmt")
params.Add "vatAmt",            request.form("plainVatAmt")
params.Add "taxFreeAmt",        request.form("plainTaxFreeAmt")
params.Add "svcAmt",            request.form("plainSvcAmt")
params.Add "clipCustNm",        request.form("plainClipCustNm")
params.Add "clipCustCi",        request.form("plainClipCustCi")
params.Add "clipCustPhoneNo",   request.form("plainClipCustPhoneNo")



'============================================================================================================================================
'   SHA256 해쉬 처리
'조합 필드 : 상점아이디 + 결제수단 + 상점주문번호 + 요청일자 + 요청시간 + 거래금액(평문) + 라이센스키
'============================================================================================================================================
Dim hashPlain : hashPlain = mchtId & method & mchtTrdNo & trdDt & trdTm & trdAmt & LICENSE_KEY
Dim hashCipher : hashCipher = ""

On error resume next
hashCipher = SHA256_Encrypt(hashPlain)'해쉬 값 계산
If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & mchtTrdNo & "][SHA256 HASHING] Hashing Fail! " & err.number & " : " & err.description  )
else
    call log_message(LOG_FILE, "[" & mchtTrdNo & "][SHA256 HASHING] Plain Text[" & hashPlain & "] ---> Cipher Text[" & hashCipher & "]")
    rsp.set "hashCipher", hashCipher ' SHA256 해쉬 결과 저장
End If
On Error GoTo 0



'============================================================================================================================================
'   AES256 암호화 처리
'============================================================================================================================================
On error resume next
Dim objCrypto : Set objCrypto = Server.CreateObject("SBCryptoUtil.CryptoUtil.1") '암복호화 객체 생성
Dim encParams : set encParams = JSON.parse("{}") '암호화 파라미터들
Dim key, value, idx
Dim aesPlain
Dim aesCipher

key = params.Keys
value = params.Items
For idx = 0 To params.Count -1
    encParams.set key(idx), ""
    aesPlain = value(idx)
    
    if "" <> aesPlain  then
        aesCipher = objCrypto.EncryptBase64(AES256_KEY, aesPlain) '암호화
        encParams.set key(idx), aesCipher '암호화된 데이터로 세팅

        call log_message(LOG_FILE, "[" & mchtTrdNo & "][AES256 Encrypt] " & key(idx) & "[" & aesPlain & "] ---> [" & aesCipher & "]")
    end if
Next

If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & mchtTrdNo & "][AES256 Encrypt] Fail! " & err.number & " : " & err.description  )
Else
    rsp.set "encParams", encParams'aes256 암호화 결과 저장
End If
On Error GoTo 0
Set objCrypto = nothing

'결과 리턴
response.write JSON.stringify(rsp)
%>
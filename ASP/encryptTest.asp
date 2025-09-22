<% @CODEPAGE="65001" language="VBScript" %>
<%
Response.CharSet="utf-8"
Session.codepage="65001"
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"
%>
<%

Dim objCrypto : Set objCrypto = Server.CreateObject("SBCryptoUtil.CryptoUtil.1") '암복호화 객체 생성(AES-256-ECB -> Base64)


Dim AES256_KEY : AES256_KEY = "pgSettle30y739r82jtd709yOfZ2yK5K" '암호화 키값
Dim plainText : plainText ="TEST" '평문
Dim cipherText '암호문


cipherText = objCrypto.encryptBase64(AES256_KEY, plainText) '암호화
response.write("PlainText["&plainText&"] ---> CipherText["&cipherText&"]<br>")

plainText = objCrypto.decryptBase64(AES256_KEY, cipherText) '복호화
response.write("CipherText["&cipherText&"] ---> PlainText["&plainText&"]<br>")


set objCrypto = nothing
%>
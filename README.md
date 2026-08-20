# ASP-PG-SDK

헥토파이낸셜 PG 연동을 위한 ASP 샘플 코드입니다.

## 📋 개요

본 샘플코드는 **표준 결제창(UI) 방식**으로 결제를 처리합니다.
- 결제창을 통한 사용자 인터페이스 제공
- 다양한 결제 수단 지원 (신용카드, 계좌이체, 가상계좌, 휴대폰 등)
- 자동결제 및 정기결제 기능 지원

**⚠️ 주의사항**: ASP 이용 가맹점의 경우 DLL 설치가이드(ASP Classic)를 참조하여 설치 바랍니다.

## 📁 파일 구조

```
/(Project Root)
├─ASP/
│  │  index.html                    <--- 인덱스 페이지
│  │  encryptTest.asp               <--- COM+ 컴포넌트 호출 테스트 페이지 (암복호화 모듈)
│  │
│  │  pay_form.asp                  <--- 결제 메인 폼
│  │  pay_encryptParams.asp         <--- 결제 파라미터 암호화 및 해시 처리 페이지
│  │  pay_autoPayResult.asp         <--- 휴대폰 자동결제 결과 페이지
│  │  pay_receiveResult.asp         <--- 결제 완료 후 응답 파라미터 수신 페이지
│  │  pay_showResult.asp            <--- 자식 페이지에서 전달된 응답 파라미터 출력
│  │
│  │  cancel_form.asp               <--- 취소 메인 폼
│  │  cancel_showResult.asp         <--- 취소 처리 및 결과 화면
│  │
│  │  receiveNoti.asp               <--- 결제 완료 후 노티 수신 페이지
│  │  processNoti.asp               <--- 노티 수신 후 처리하는 페이지
│  │
│  └─inc/
│          config.asp               <--- 기본 정보 설정 파일 (*자사에 맞게 변경 필요)
│          json2.asp                <--- JSON 라이브러리
│          KISA_SHA256.asp          <--- KISA에서 배포한 SHA256 라이브러리
│          settleUtils.asp          <--- 헥토파이낸셜 유틸 라이브러리
│
└─DLL/
        libiconv.dll                <--- SBCryptoUtil.dll에서 의존하는 iconv dll
        libiconvD.dll               <--- SBCryptoUtil.dll에서 의존하는 iconv dll
        SBCryptoUtil.dll            <--- AES256 암호화 동적 라이브러리
```

## 📄 파일 설명

### 🔧 공통 페이지
- **index.html**: 인덱스 페이지입니다.
- **encryptTest.asp**: 암복호화 컴포넌트 호출 테스트 페이지입니다.
- **config.asp**: 상점아이디, 암복호화키 등을 설정할 수 있는 설정 파일입니다.
- **receiveNoti.asp**: 결제 또는 취소 처리가 완료된 후, 헥토파이낸셜에서 가맹점으로 전달하는 노티(결과통보)를 수신하는 페이지입니다.
- **processNoti.asp**: receiveNoti.asp에서 결제 또는 취소의 성공/실패에 따라 적절한 로직을 수행하는 메소드를 정의한 파일입니다.

### 💳 결제 관련 페이지
- **pay_form.asp**: 결제 요청 시 사용자로부터 정보를 입력받는 Form 페이지입니다. 결제는 Form POST 방식으로 처리됩니다.
- **pay_encryptParams.asp**: pay_form.asp에서 암호화가 필요한 파라미터들을 AJAX 통신으로 암호화하는 페이지입니다. 또한 SHA256 해시 처리도 수행합니다.
- **pay_receiveResult.asp**: 결제창에서 결제가 완료된 이후 닫기 버튼을 누를 때, 헥토파이낸셜로부터 응답 파라미터를 수신하는 페이지입니다.
- **pay_showResult.asp**: pay_receiveResult.asp에서 받은 파라미터를 부모창으로 전송할 수 있는데, 이때 전송된 파라미터들을 수신하여 출력하는 페이지입니다.
- **pay_autoPayResult.asp**: 휴대폰 자동연장결제 시 사용되는 결제 및 결과 페이지입니다.

### ❌ 취소 관련 페이지
- **cancel_form.asp**: 취소 요청 시 사용자로부터 정보를 입력받는 Form 페이지입니다.
- **cancel_showResult.asp**: 헥토파이낸셜과 Server to Server로 커넥션하여, 취소 요청을 하고 응답을 받아 결과를 출력하는 페이지입니다.

## 🔄 프로세스 처리 순서

- **결제 처리 순서**: pay_form.asp → pay_encryptParams.asp → pay_receiveResult.asp → pay_showResult.asp
- **휴대폰 자동연장 결제**: pay_form.asp → pay_autoPayResult.asp
- **취소 처리 순서**: cancel_form.asp → cancel_showResult.asp
- **노티 처리 순서**: receiveNoti.asp → processNoti.asp

## ⚙️ config.asp 설정 파일 변수 설명

- **PG_MID**: 상점아이디. 테스트환경에서의 상점아이디는 샘플소스에 기재되어 있습니다. 상용테스트 시에는 헥토파이낸셜에서 발급한 MID로 설정하셔야 합니다. 이 값은 외부에 노출되어서는 안됩니다.
- **LICENSE_KEY**: MID당 하나의 라이센스키가 발급됩니다. SHA256 해시체크 용도로 사용됩니다. 이 값은 외부에 노출되어서는 안됩니다.
- **AES256_KEY**: 개인정보/민감정보를 암복호화하는데 사용되는 키로서, 외부에 노출되어서는 안됩니다.
- **PAYMENT_SERVER**: 헥토파이낸셜 결제 처리 서버의 URL입니다. 변경하지 마십시오.
- **CANCEL_SERVER**: 헥토파이낸셜 취소 처리 서버의 URL입니다. 변경하지 마십시오.
- **CONN_TIMEOUT**: 헥토파이낸셜 API 통신 연결 타임아웃입니다.
- **READ_TIMEOUT**: 헥토파이낸셜 API 통신 수신 타임아웃입니다.

## 📢 노티 수신 페이지

- **파일명**: receiveNoti.asp
- 결제 또는 취소 완료 후 헥토파이낸셜 서버에서 콜백으로 호출하게 되는 페이지이며, 헥토파이낸셜에서 가맹점으로 노티를 전송합니다.
- nextUrl(결과페이지)에서는 성공/실패에 대한 결과 화면을 고객에게 리턴하여 주시고,
- notiUrl(노티수신페이지)에서는 가맹점의 실제 내부데이터, DB를 처리하시면 됩니다.

## 문의

- 기술 문의: pgsupport@hecto.co.kr
- 개발 가이드: [헥토파이낸셜 개발자 센터](https://developers.hectofinancial.co.kr)

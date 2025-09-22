# ASP-PG-SDK

헥토파이낸셜 PG 연동을 위한 ASP 샘플 코드입니다.

## 📋 목차

- [개요](#개요)
- [파일 구조](#파일-구조)
- [설치 및 설정](#설치-및-설정)
- [사용법](#사용법)
- [API 문서](#api-문서)
- [라이센스](#라이센스)

## 🚀 개요

이 SDK는 헥토파이낸셜 PG 서비스와 연동하기 위한 ASP(Active Server Pages) 샘플 코드를 제공합니다. 
결제, 취소, 노티 수신 등의 기능을 포함하고 있습니다.

## 📁 파일 구조

```
PG2_ASP_v1.6/
├── ASP/                          # ASP 소스 파일
│   ├── index.html               # 인덱스 페이지
│   ├── encryptTest.asp          # 암복호화 테스트 페이지
│   ├── pay_form.asp             # 결제 폼
│   ├── pay_encryptParams.asp    # 결제 파라미터 암호화
│   ├── pay_autoPayResult.asp    # 휴대폰 자동결제 결과
│   ├── pay_receiveResult.asp    # 결제 결과 수신
│   ├── pay_showResult.asp       # 결제 결과 출력
│   ├── cancel_form.asp          # 취소 폼
│   ├── cancel_showResult.asp    # 취소 결과
│   ├── receiveNoti.asp          # 노티 수신
│   ├── processNoti.asp          # 노티 처리
│   └── inc/                     # 공통 라이브러리
│       ├── config.asp           # 설정 파일
│       ├── json2.asp            # JSON 라이브러리
│       ├── KISA_SHA256.asp      # SHA256 라이브러리
│       └── settleUtils.asp      # 유틸리티 함수
└── DLL/                         # 암호화 DLL
    ├── libiconv.dll
    ├── libiconvD.dll
    └── SBCryptoUtil.dll
```

## ⚙️ 설치 및 설정

### 1. 환경 요구사항

- Windows Server with IIS
- ASP 지원 환경
- COM+ 컴포넌트 등록 가능한 환경

### 2. 설정 파일 수정

`ASP/inc/config.asp` 파일에서 다음 값들을 설정하세요:

```asp
' 상점 정보
PG_MID = "your_merchant_id"           ' 상점아이디
LICENSE_KEY = "your_license_key"      ' 라이센스키
AES256_KEY = "your_aes256_key"        ' AES256 암호화키

' 서버 URL (변경하지 마세요)
PAYMENT_SERVER = "https://pg.hectofinancial.com/payment"
CANCEL_SERVER = "https://pg.hectofinancial.com/cancel"

' 타임아웃 설정
CONN_TIMEOUT = 30
READ_TIMEOUT = 30
```

### 3. DLL 등록

관리자 권한으로 명령 프롬프트를 실행하고 다음 명령을 실행하세요:

```cmd
regsvr32 SBCryptoUtil.dll
```

## 🔧 사용법

### 결제 처리

1. **결제 폼**: `pay_form.asp`에서 결제 정보 입력
2. **파라미터 암호화**: `pay_encryptParams.asp`에서 민감정보 암호화
3. **결제 처리**: 세틀뱅크 서버로 결제 요청
4. **결과 수신**: `pay_receiveResult.asp`에서 결과 수신
5. **결과 출력**: `pay_showResult.asp`에서 결과 화면 출력

### 취소 처리

1. **취소 폼**: `cancel_form.asp`에서 취소 정보 입력
2. **취소 처리**: `cancel_showResult.asp`에서 취소 요청 및 결과 출력

### 노티 처리

1. **노티 수신**: `receiveNoti.asp`에서 세틀뱅크로부터 노티 수신
2. **노티 처리**: `processNoti.asp`에서 비즈니스 로직 처리

## 📚 API 문서

### 주요 함수

#### SettleUtils.asp

- `encryptData(data)`: 데이터 AES256 암호화
- `decryptData(data)`: 데이터 AES256 복호화
- `makeHash(params)`: SHA256 해시 생성
- `sendHttpRequest(url, params)`: HTTP 요청 전송

### 프로세스 플로우

```
결제: pay_form.asp → pay_encryptParams.asp → pay_receiveResult.asp → pay_showResult.asp
취소: cancel_form.asp → cancel_showResult.asp
노티: receiveNoti.asp → processNoti.asp
```

## 🔒 보안 주의사항

- **중요**: `config.asp`의 민감한 정보(PG_MID, LICENSE_KEY, AES256_KEY)는 외부에 노출되지 않도록 주의하세요.
- 프로덕션 환경에서는 HTTPS를 사용하세요.
- 정기적으로 암호화 키를 변경하세요.

## 🐛 문제 해결

### 자주 발생하는 문제

1. **DLL 등록 오류**: 관리자 권한으로 실행하세요.
2. **암호화 오류**: AES256_KEY가 올바른지 확인하세요.
3. **통신 오류**: 방화벽 설정을 확인하세요.

## 📞 지원

- 기술 지원: [헥토파이낸셜 고객센터]
- 문서: [헥토파이낸셜 개발자 문서]

## 📄 라이센스

이 프로젝트는 헥토파이낸셜의 라이센스 하에 제공됩니다.

---

**버전**: v1.6  
**최종 업데이트**: 2024년 9월

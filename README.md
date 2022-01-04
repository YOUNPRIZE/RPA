<img src="https://img.shields.io/badge/Python-3776AB?style=flat&logo=Python&logoColor=white"/> <img src="https://img.shields.io/badge/Google Drive-4285F4?style=flat&logo=Google Drive&logoColor=white"/> <img src="https://img.shields.io/badge/Google Sheets-34A853?style=flat&logo=Google Sheets&logoColor=white"/>

# RPA (Robotic Process Automation)
## Outline
다양한 OPEN API를 사용할 수 있는 공공데이터포털(https://www.data.go.kr) 에서 주유소의 요소수 재고 현황을 제공해주는 API를 사용해 SKT의 인공지능 스피커인 NUGU를 통해 요소수의 정보를 알려주는 기능입니다.
## Description
> **Structure**
> - User의 발화 -> NUGU -> Backend Proxy -> NUGU -> User
>
> ![main](./png/search.png)
>
> 1. User가 '(지명)에 요소수 파는 곳 알려줘' 라고 발화
> 
> 2. NUGU에서 (지명)을 인식
>
> 3. Backend Proxy 에서 (지명)에 위치하고 있는 주유소의 정보(주소, 전화번호) 추출
>
> 4. NUGU가 "(지명)에 있는 주유소의 전화번호는 (번호), 주소는 (주소)입니다."라는 발화 Output
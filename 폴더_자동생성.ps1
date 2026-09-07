# 젠트리 자원관리 작업대 — 폴더 자동 생성 스크립트
# 사용법: PowerShell에서 이 파일을 실행하세요
#        .\폴더_자동생성.ps1

$basePath = "C:\Users\SUNJIN\Desktop\작업대"
$timestamp = Get-Date -Format "yyyy-MM-dd HHmm"

Write-Host "🔄 젠트리 자원관리 작업대 폴더 구조 생성 시작..." -ForegroundColor Cyan
Write-Host "기본 경로: $basePath`n" -ForegroundColor Gray

# 폴더 생성 함수
function New-FolderStructure {
    param([string]$Path, [string]$Name)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Host "✅ 생성됨: $Name" -ForegroundColor Green
    } else {
        Write-Host "⏭️  이미 있음: $Name" -ForegroundColor Yellow
    }
}

# 파일 생성 함수 (덮어쓰지 않음)
function New-FileIfNotExists {
    param([string]$Path, [string]$Content)
    if (-not (Test-Path $Path)) {
        Set-Content -Path $Path -Value $Content -Encoding UTF8
        Write-Host "✅ 파일 생성: $(Split-Path -Leaf $Path)" -ForegroundColor Green
    } else {
        Write-Host "⏭️  파일 존재: $(Split-Path -Leaf $Path)" -ForegroundColor Yellow
    }
}

# ========== 젠트리 폴더 ==========
Write-Host "`n📁 [젠트리] 개인 자원 폴더 생성..." -ForegroundColor Magenta

New-FolderStructure "$basePath\젠트리" "젠트리 (기본 폴더)"
New-FolderStructure "$basePath\젠트리\목공방" "목공방 (훈련)"
New-FolderStructure "$basePath\젠트리\목공방\작품" "작품 (이미지, 설명)"
New-FolderStructure "$basePath\젠트리\월1000" "월1000 (재테크)"
New-FolderStructure "$basePath\젠트리\월1000\뉴스" "뉴스 (종목별 클리핑)"
New-FolderStructure "$basePath\젠트리\월1000\분석" "분석 (일일 브리핑)"
New-FolderStructure "$basePath\젠트리\취미" "취미 (낚시, 텃밭, 골프)"
New-FolderStructure "$basePath\젠트리\자원대시보드" "자원대시보드 (HTML)"
New-FolderStructure "$basePath\젠트리\자원대시보드\백업" "HTML 백업"

# ========== 가족 폴더 ==========
Write-Host "`n👨‍👩‍👧‍👦 [가족] 자원 폴더 생성..." -ForegroundColor Magenta

New-FolderStructure "$basePath\가족" "가족 (기본 폴더)"
New-FolderStructure "$basePath\가족\딸" "딸 (의사)"
New-FolderStructure "$basePath\가족\아들" "아들 (군→복학→취업)"
New-FolderStructure "$basePath\가족\배우자" "배우자"

# ========== 초기 마크다운 파일 생성 ==========
Write-Host "`n📝 초기 문서 생성..." -ForegroundColor Magenta

# 젠트리 - 목공방
New-FileIfNotExists "$basePath\젠트리\목공방\진도.md" @"
# 목공방 훈련 진도

## 목표
제주에 목공방(공방) 개설 → 제2의 인생 시작

## 현재 상태
- 교육: AX 건축목공반 2기
- 진도: __ / __회차 (직접 입력)
- 수료 예정: 2026년 __ 월

## 남은 과제
- [ ] 과제 1
- [ ] 과제 2
- [ ] 수료작

## 완성 작품
(작품/ 폴더에 이미지와 함께 정리)

---
마지막 수정: $(Get-Date -Format 'yyyy-MM-dd')
"@

New-FileIfNotExists "$basePath\젠트리\목공방\자격시험.md" @"
# 자격시험 준비

## 목표
건축목공기능사 자격 취득

## 시험 일정
| 시험 | 상태 | 시험일 | D-day |
|------|------|--------|--------|
| 필기 | __ | ____ | __ |
| 실기 | __ | ____ | __ |

## 준비 현황
- 필기:
- 실기:

---
마지막 수정: $(Get-Date -Format 'yyyy-MM-dd')
"@

New-FileIfNotExists "$basePath\젠트리\목공방\일정.md" @"
# 훈련 일정

## 이번 주 (__ 월 __ 일 ~ __ 월 __ 일)
- [ ] __ 월 __ 일 (목):
- [ ] __ 월 __ 일 (금):
- [ ] 토/일: 자기공부

## 다가올 중요 날짜
- 시험 일정
- 수료 예정일
- 기타 공지

---
마지막 수정: $(Get-Date -Format 'yyyy-MM-dd')
"@

# 월1000
New-FileIfNotExists "$basePath\젠트리\월1000\포트폴리오.md" @"
# 월1000 포트폴리오 현황

## 목표
월 1,000만원 배당 수익 → 평생 경제적 자유

## 보유 종목 (4개)

| 종목명 | 티커 | 시장 | 계좌 | 현황 |
|--------|------|------|------|------|
| RISE 코리아밸류업위클리고정커버드콜 | 0094M0 | ETF | IRP/ISA | __ |
| SOL 팔란티어커버드콜OTM채권혼합 | 0040Y0 | ETF | IRP/ISA | 월분배 230원 |
| RISE 미국배당100데일리고정커버드콜 | 490600 | ETF | IRP/ISA | __ |
| 비에이치아이(BHI) | 083650 | 코스닥 | IRP/ISA | __ |

## 이번 달 현황
- 총 자산: __원
- 예상 월배당: __원
- 목표까지: __원 (D-day: __)

## 자동화 진행
- ✅ pykrx 설치 완료
- ✅ 일일 브리핑 스케줄 등록 (매일 08:00)
- ⏳ PlayMCP(카톡) 연결 진행 중
- 📧 Gmail 브리핑 발송 (매일)

---
마지막 수정: $(Get-Date -Format 'yyyy-MM-dd')
"@

# 가족
New-FileIfNotExists "$basePath\가족\가족현황.md" @"
# 가족 자원 현황

## 가족 구성
| 이름 | 나이 | 직업/상태 | 현재 초점 |
|------|------|----------|----------|
| 젠트리(나) | __ | 목공 훈련 + 재테크 | 훈련, 월1000 |
| 배우자 | 53 | __ | __ |
| 딸 | 25 | 의사 | 진료, 경력 |
| 아들 | 21 | 군복무 중 | 군, 복학 예정 |

## 중요 일정
- [ ] 아들 제대: __년 __월 (D-__)
- [ ] 아들 복학: __년 __월
- [ ] 아들 졸업: __년 __월
- [ ] 딸 (추적할 일정): __

## 가족 자원 추적
- 시간 배분: (누가 뭘 하는지)
- 재정: (수입/지출/지원)
- 장기 계획: (5년 후, 10년 후)

---
마지막 수정: $(Get-Date -Format 'yyyy-MM-dd')
"@

New-FileIfNotExists "$basePath\가족\아들\일정.md" @"
# 아들 일정 관리

## 군 복무
- 입대일: __년 __월 __ 일
- 제대예정: __년 __월 __ 일 (D-__)
- 휴가 계획:

## 복학 (제대 후)
- 복학 예정일: __년 __월
- 복학할 학과: __(__)
- 졸업 예정일: __년 __월

## 취업 준비
- 취업 목표:
- 지원 회사:
- 인터뷰 일정:
- 필요 준비물:

---
마지막 수정: $(Get-Date -Format 'yyyy-MM-dd')
"@

Write-Host "`n✨ 모든 폴더와 초기 문서 생성 완료!" -ForegroundColor Cyan
Write-Host "다음 단계:`n  1. 각 .md 파일을 열어서 내용 입력`n  2. 작업대 HTML 커스터마이즈`n  3. 인덱스.md 작성" -ForegroundColor Yellow

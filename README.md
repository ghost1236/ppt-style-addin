# PowerPoint 디자인 스타일 Add-in

PowerPoint에 커스텀 리본 메뉴와 사이드패널을 추가하는 Office JS Task Pane Add-in.
전체 슬라이드 또는 선택 영역에 폰트/컬러/사이즈 스타일을 일괄 적용하는 디자인 도구.

**Windows / Mac / 웹 PowerPoint (Microsoft 365) 모두 동작.**

## 기능

- **스타일 프리셋 관리**: 자주 쓰는 스타일 조합을 저장/불러오기/삭제
- **적용 대상 선택**: 선택한 텍스트, 현재 슬라이드, 전체 슬라이드 등 8가지 옵션
- **즉시 스타일 적용**: 폰트명/크기/스타일/색상/정렬/줄간격 실시간 편집
- **JSON 내보내기/불러오기**: 프리셋을 파일로 공유
- **실행 취소**: 최대 10단계
- **버전 자동 감지**: Microsoft 365 / 영구 라이선스 환경 자동 구분 및 fallback 처리

## 설치 및 실행

```bash
# 1. 의존성 설치
npm install

# 2. HTTPS 인증서 생성 (최초 1회)
npx office-addin-dev-certs install

# 3. 개발 서버 실행
npm run dev

# 4. PowerPoint에서 로드
#    삽입 → 내 추가 기능 → 공유 폴더에서 manifest.xml 등록

# 5. 프로덕션 빌드
npm run build
```

## 팀 배포 (공유 폴더 방식)

1. 네트워크 공유 폴더에 빌드 결과물(`dist/`) + `manifest.xml` 배치
2. PowerPoint → 파일 → 옵션 → 보안 센터 → 신뢰할 수 있는 추가 기능 카탈로그 등록
3. 팀원: `삽입 → 내 추가 기능 → 내 조직`에서 설치

## 지원 환경

| 환경 | 지원 여부 |
|------|----------|
| Windows PowerPoint (Microsoft 365) | ✅ 완전 지원 |
| Mac PowerPoint (Microsoft 365) | ✅ 완전 지원 |
| 웹 PowerPoint (office.com) | ✅ 완전 지원 |
| Windows PowerPoint 2019/2021 영구 라이선스 | ⚠️ 일부 API 제한 |
| Mac PowerPoint 2019/2021 영구 라이선스 | ⚠️ 일부 API 제한 |
| Mac PowerPoint 2016 이하 | ❌ 미지원 |

## 프로젝트 구조

```
PPTAdon/
├── manifest.xml              # Office Add-in 매니페스트
├── assets/                   # 아이콘 (SVG)
├── src/
│   ├── taskpane/
│   │   ├── index.html
│   │   ├── main.tsx
│   │   └── components/
│   │       ├── App.tsx             # 루트 컴포넌트
│   │       ├── TargetSelector.tsx  # 적용 대상 선택
│   │       ├── StyleEditor.tsx     # 폰트/컬러/사이즈 편집
│   │       ├── PresetList.tsx      # 프리셋 목록
│   │       ├── PresetCard.tsx      # 개별 프리셋 카드
│   │       ├── PresetModal.tsx     # 저장/수정 모달
│   │       └── ColorPicker.tsx     # react-colorful 래퍼
│   ├── commands/
│   │   ├── commands.html           # 리본 버튼 진입점
│   │   └── commands.ts             # 리본 버튼 실행 로직
│   ├── services/
│   │   ├── officeService.ts        # Office.js API 래퍼 및 버전 감지
│   │   ├── styleService.ts         # 스타일 적용 로직
│   │   └── presetStorage.ts        # 프리셋 저장/불러오기
│   └── store/
│       └── useStore.ts             # Zustand 상태 관리
├── vite.config.ts
├── tsconfig.json
└── package.json
```

## 아이콘 교체

`assets/` 폴더의 SVG 파일을 PNG로 교체하고 `manifest.xml`의 URL을 업데이트하세요.
- `icon-16.png` (16×16px)
- `icon-32.png` (32×32px)
- `icon-80.png` (80×80px)

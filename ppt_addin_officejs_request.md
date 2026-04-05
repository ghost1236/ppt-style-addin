# PowerPoint 디자인 스타일 Add-in 개발 요청서 (Office JS 방식)

## 프로젝트 개요

PowerPoint에 커스텀 리본 메뉴와 사이드패널을 추가하는 Office JS Task Pane Add-in.
전체 슬라이드 또는 선택 영역에 폰트/컬러/사이즈 스타일을 일괄 적용하는 디자인 도구.
**Windows / Mac / 웹 PowerPoint (Microsoft 365) 모두 동작.**

---

## 기술 스택

- **Add-in 타입**: Office JS Task Pane Add-in (사이드패널)
- **Framework**: React + Vite + TypeScript
- **Office API**: Office.js (`@types/office-js`)
- **UI**: Fluent UI React v9 (Microsoft 공식 디자인 시스템)
- **상태 관리**: Zustand
- **색상 피커**: `react-colorful`
- **설정 저장**: `Office.context.document.settings` (Add-in 전용 저장소)
- **번들러**: Vite
- **인증서**: office-addin-dev-certs (로컬 개발용 HTTPS)
- **스캐폴딩**: Yeoman + generator-office (초기 프로젝트 생성)

---

## 핵심 기능

---

### 기능 1. 스타일 프리셋 관리

사이드패널에서 자주 쓰는 스타일 조합을 프리셋으로 저장하고 불러오는 기능.

#### 프리셋 구성 항목
```typescript
interface StylePreset {
  id: string;
  name: string;           // 예: "제목 스타일", "강조 텍스트"
  font: {
    name?: string;        // 예: "Pretendard", "Malgun Gothic"
    size?: number;        // pt 단위
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;       // HEX (예: "#FF0000")
  };
  paragraph?: {
    alignment?: 'left' | 'center' | 'right' | 'justify';
    lineSpacing?: number; // 줄간격 (%)
  };
}
```

#### 저장 방식
- `Office.context.document.settings`에 JSON으로 저장
- 같은 파일을 열면 프리셋 유지
- 내보내기/불러오기: JSON 파일로 export/import

---

### 기능 2. 스타일 적용 대상 선택

사이드패널 상단에서 적용 대상을 선택:

```
적용 대상:
○ 선택한 텍스트만        → 드래그로 선택한 텍스트 범위에만 적용
○ 선택한 텍스트 상자     → 클릭한 도형/텍스트박스 전체
○ 현재 슬라이드 - 제목   → 현재 슬라이드의 제목 placeholder만
○ 현재 슬라이드 - 본문   → 현재 슬라이드의 본문 placeholder만
○ 현재 슬라이드 - 전체   → 현재 슬라이드 모든 텍스트
○ 모든 슬라이드 - 제목   → 전체 슬라이드 제목 일괄 적용
○ 모든 슬라이드 - 본문   → 전체 슬라이드 본문 일괄 적용
○ 모든 슬라이드 - 전체   → 전체 슬라이드 모든 텍스트 일괄 적용
```

---

### 기능 3. 스타일 즉시 적용 패널

프리셋 없이 바로 개별 속성을 조정하고 적용하는 패널.

#### UI 구성
```
── 폰트 ──────────────────────────
폰트명    [Pretendard          ▼]   ← 시스템 폰트 목록 드롭다운
크기      [24    ] pt
스타일    [B] [I] [U]             ← Bold / Italic / Underline 토글
색상      [████] #333333          ← 컬러 피커

── 단락 ──────────────────────────
정렬      [◀] [≡] [▶] [≣]        ← 좌/중/우/양쪽
줄간격    [1.5  ] 배

── 적용 ──────────────────────────
[현재 선택에 적용]  [프리셋으로 저장]
```

---

### 기능 4. 프리셋 목록 패널

```
── 저장된 프리셋 ──────────────────
┌─────────────────────────────────┐
│ 🎨 제목 스타일                  │
│ Pretendard Bold 36pt #1A1A2E   │
│              [적용] [수정] [삭제]│
├─────────────────────────────────┤
│ 🎨 본문 스타일                  │
│ Malgun Gothic 18pt #444444     │
│              [적용] [수정] [삭제]│
├─────────────────────────────────┤
│ 🎨 강조 텍스트                  │
│ Pretendard Bold 20pt #E63946   │
│              [적용] [수정] [삭제]│
└─────────────────────────────────┘
[+ 새 프리셋]  [JSON 내보내기] [불러오기]
```

---

### 기능 5. 실행 취소 지원

- 적용 전 상태를 스냅샷으로 저장
- 사이드패널 내 `[↩ 실행취소]` 버튼
- 최대 10단계 유지

---

## 리본 메뉴 구성

```
[홈] [삽입] [디자인] ... [디자인 도구]
                              └ 스타일 패널 열기
                              └ 제목 일괄 적용
                              └ 본문 일괄 적용
```

---

## Office.js API 활용 포인트

### 슬라이드 전체 순회 + 스타일 적용
```typescript
await PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load('items');
  await context.sync();

  for (const slide of slides.items) {
    const shapes = slide.shapes;
    shapes.load('items');
    await context.sync();

    for (const shape of shapes.items) {
      if (shape.textFrame) {
        const tf = shape.textFrame;
        tf.textRange.font.name = preset.font.name;
        tf.textRange.font.size = preset.font.size;
        tf.textRange.font.color = preset.font.color;
        tf.textRange.font.bold = preset.font.bold;
      }
    }
  }
  await context.sync();
});
```

### Placeholder 타입 구분 (제목 vs 본문)
```typescript
// shape.placeholderType으로 구분
// PowerPoint.PlaceholderType.title → 제목
// PowerPoint.PlaceholderType.body  → 본문
```

---

## 프로젝트 구조

```
ppt-style-addin/
├── manifest.xml
├── assets/                         # 아이콘 (16/32/80px PNG)
├── src/
│   ├── taskpane/
│   │   ├── index.html
│   │   ├── main.tsx
│   │   └── components/
│   │       ├── App.tsx
│   │       ├── TargetSelector.tsx  # 적용 대상 선택
│   │       ├── StyleEditor.tsx     # 폰트/컬러/사이즈 입력
│   │       ├── PresetList.tsx      # 프리셋 목록
│   │       ├── PresetCard.tsx      # 개별 프리셋 카드
│   │       ├── PresetModal.tsx     # 저장/수정 모달
│   │       └── ColorPicker.tsx     # react-colorful 래퍼
│   ├── commands/
│   │   ├── commands.html           # 리본 버튼 진입점
│   │   └── commands.ts             # 리본 버튼 실행 로직
│   ├── services/
│   │   ├── officeService.ts        # Office.js API 래퍼
│   │   ├── styleService.ts         # 스타일 적용 로직
│   │   └── presetStorage.ts        # 프리셋 저장/불러오기
│   └── store/
│       └── useStore.ts
├── vite.config.ts
└── package.json
```

---

## 로컬 개발 및 설치 방법 (README에 포함)

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

### 팀 배포 (공유 폴더 방식)
- 네트워크 공유 폴더에 빌드 결과물 + `manifest.xml` 배치
- PowerPoint → 파일 → 옵션 → 보안 센터 → 신뢰할 수 있는 추가 기능 카탈로그 등록
- 팀원: `삽입 → 내 추가 기능 → 내 조직`에서 설치

---

## 지원 환경

| 환경 | 지원 여부 |
|------|----------|
| Windows PowerPoint (Microsoft 365) | ✅ 완전 지원 |
| Mac PowerPoint (Microsoft 365) | ✅ 완전 지원 |
| 웹 PowerPoint (office.com) | ✅ 완전 지원 |
| Windows PowerPoint 2019/2021 영구 라이선스 | ⚠️ 일부 API 제한 — 하단 대응 전략 참고 |
| Mac PowerPoint 2019/2021 영구 라이선스 | ⚠️ 일부 API 제한 — 하단 대응 전략 참고 |
| Mac PowerPoint 2016 이하 | ❌ 미지원 |

---

## 버전별 대응 전략 (중요)

사용자 환경이 Microsoft 365인지 영구 라이선스인지 알 수 없으므로,
**런타임에 API 가용 여부를 체크하고 자동으로 fallback 처리**하도록 구현할 것.

### API 버전 체크 방식
```typescript
// 실행 전 API 지원 여부 확인
const isApiSupported = (requirement: string, version: string): boolean => {
  return Office.context.requirements.isSetSupported(requirement, version);
};

// 예: PresentationAPI 1.3 이상 필요한 기능
if (isApiSupported('PresentationAPI', '1.3')) {
  // Microsoft 365 최신 API 사용
  await applyWithNewApi(preset);
} else {
  // 영구 라이선스 fallback — 구버전 API로 처리
  await applyWithLegacyApi(preset);
}
```

### 기능별 버전 분기표

| 기능 | Microsoft 365 | 영구 라이선스 2019/2021 | Fallback 방법 |
|------|:---:|:---:|------|
| 슬라이드 순회 | PresentationAPI 1.2 | ✅ 동일 지원 | — |
| 폰트/사이즈/색상 적용 | PresentationAPI 1.2 | ✅ 동일 지원 | — |
| Placeholder 타입 구분 | PresentationAPI 1.3 | ⚠️ 미지원 | shape 이름/인덱스로 추정 |
| 선택 텍스트 범위 접근 | Office.js 공통 | ✅ 동일 지원 | — |
| 줄간격(lineSpacing) | PresentationAPI 1.5 | ⚠️ 미지원 | 적용 불가 안내 토스트 표시 |

### Placeholder 타입 구분 Fallback (영구 라이선스용)
```typescript
// PresentationAPI 1.3 미지원 시: shape 이름으로 제목/본문 추정
const isTitleShape = (shapeName: string): boolean => {
  const titleKeywords = ['Title', '제목', 'title'];
  return titleKeywords.some(k => shapeName.includes(k));
};

const isBodyShape = (shapeName: string): boolean => {
  const bodyKeywords = ['Content', 'Body', '내용', '본문', 'Text'];
  return bodyKeywords.some(k => shapeName.includes(k));
};
```

### UI 처리 방침
- 영구 라이선스 환경에서 지원되지 않는 기능은 **버튼을 비활성화(disabled)**
- 비활성화된 버튼에 호버 시 툴팁: `"이 기능은 Microsoft 365에서만 지원됩니다"`
- 사이드패널 하단에 현재 감지된 Office 버전 표시:
  ```
  ℹ️ 감지된 환경: PowerPoint 2021 (일부 기능 제한)
  ```

### manifest.xml Requirements 설정
```xml
<!-- 최소 요구사항: PresentationAPI 1.2 (영구 라이선스도 지원) -->
<Requirements>
  <Sets DefaultMinVersion="1.1">
    <Set Name="PresentationAPI" MinVersion="1.2"/>
  </Sets>
</Requirements>
```
MinVersion을 1.2로 낮춰 영구 라이선스에서도 Add-in 자체는 로드되게 하고,
고급 기능만 런타임 체크로 분기 처리.

---

## 개발 우선순위

1. 프로젝트 초기 세팅 (Yeoman + Vite + React 전환)
2. 사이드패널 기본 레이아웃 (Fluent UI)
3. 적용 대상 선택 + 즉시 스타일 적용
4. 프리셋 저장/불러오기/적용/삭제
5. 전체 슬라이드 일괄 적용 (제목/본문/전체)
6. 실행 취소
7. 리본 메뉴 버튼 추가
8. JSON 내보내기/불러오기
9. 아이콘 및 UI 다듬기

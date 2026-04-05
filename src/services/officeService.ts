/* global Office, PowerPoint */

/** PowerPoint 전역 객체가 초기화됐는지 확인 */
export function isPowerPointReady(): boolean {
  return typeof PowerPoint !== 'undefined';
}

/**
 * API 지원 여부 확인
 */
export function isApiSupported(requirement: string, version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported(requirement, version);
  } catch {
    return false;
  }
}

/**
 * PresentationAPI 1.3 이상 (제목/본문 Placeholder 타입 구분)
 */
export const hasPlaceholderTypeSupport = (): boolean =>
  isApiSupported('PresentationAPI', '1.3');

/**
 * PresentationAPI 1.5 이상 (줄간격 lineSpacing)
 */
export const hasLineSpacingSupport = (): boolean =>
  isApiSupported('PresentationAPI', '1.5');

/**
 * shape 이름으로 제목인지 추정 (Fallback, 영구 라이선스용)
 */
export function isTitleShape(shapeName: string): boolean {
  const keywords = ['title', 'Title', '제목'];
  return keywords.some((k) => shapeName.includes(k));
}

/**
 * shape 이름으로 본문인지 추정 (Fallback, 영구 라이선스용)
 */
export function isBodyShape(shapeName: string): boolean {
  const keywords = ['content', 'Content', 'body', 'Body', '내용', '본문', 'Text'];
  return keywords.some((k) => shapeName.includes(k));
}

/**
 * Office 버전 문자열 반환
 */
export function getOfficeVersion(): string {
  try {
    const info = (Office.context as unknown as { diagnostics?: { version?: string } }).diagnostics;
    return info?.version ?? 'Unknown';
  } catch {
    return 'Unknown';
  }
}

/**
 * 현재 PowerPoint 선택 정보 가져오기 (텍스트 선택 여부)
 */
export async function getSelectionType(): Promise<string> {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve('text');
        } else {
          resolve('none');
        }
      }
    );
  });
}

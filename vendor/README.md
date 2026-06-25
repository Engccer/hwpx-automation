# vendor/

한컴 COM 자동화 시 "파일 접근 허용" 보안 대화상자를 우회하기 위한 서드파티 바이너리를 배치하는 폴더.

## 라이선스 고지

이 폴더에 들어가는 바이너리는 **본 저장소의 MIT License 적용 대상이 아닙니다.** 각 바이너리의 원 제공자 라이선스를 따릅니다. `.gitignore`에 의해 Git에는 포함되지 않으며, 사용자가 로컬에서 직접 설치해야 합니다.

## FilePathCheckerModuleExample.dll

- **제공자**: Hancom Inc. (한컴디벨로퍼 공식 샘플)
- **용도**: `hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')` 호출 시 로드되어, 파일 열기/저장 시 뜨는 보안 대화상자를 자동 승인

### 설치 절차 (Windows)

1. 한컴디벨로퍼 공식 배포처에서 DLL 다운로드
   - 주소: <https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/보안모듈(Automation).zip>
   - 압축 해제 후 `FilePathCheckerModuleExample.dll` 추출
2. 본 폴더(`vendor/`)에 DLL 배치
3. 레지스트리 등록 (관리자 권한 불필요, `HKCU`):

   ```powershell
   Set-ItemProperty -Path "HKCU:\SOFTWARE\HNC\HwpAutomation\Modules" `
     -Name "FilePathCheckerModule" `
     -Value "<이 저장소 절대경로>\vendor\FilePathCheckerModuleExample.dll"
   ```

4. Python 코드에서 모듈 등록:

   ```python
   hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
   ```

### 주의

- DLL 경로가 바뀌면 레지스트리 값도 함께 갱신해야 합니다.
- 본 저장소의 MIT License는 소스 코드에만 적용됩니다. 이 DLL은 Hancom의 배포 조건을 따릅니다.

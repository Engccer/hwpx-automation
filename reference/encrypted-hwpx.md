# 암호화된 HWPX 처리

비밀번호가 설정된 HWPX는 AES-256-CBC로 `Contents/section0.xml`, `Contents/header.xml`, `settings.xml`, 이미지 파일이 모두 암호화된다. `hwpx_edit.py` 같은 XML 직접 파서는 실패하고, python-hwpx도 마찬가지다. 한컴 COM으로 먼저 암호를 제거해야 한다.

## 증상

- `hwpx_edit.py --to-md` 실행 시 `lxml.etree.XMLSyntaxError: Start tag expected, '<' not found`
- `unzip -p file.hwpx Contents/section0.xml` 출력이 바이너리 덤프
- `META-INF/manifest.xml`에 `<odf:file-entry>` 마다 `<odf:encryption-data>` 하위 요소 존재

## 감지

```python
import zipfile

def is_encrypted_hwpx(filepath):
    with zipfile.ZipFile(filepath) as zf:
        if 'META-INF/manifest.xml' not in zf.namelist():
            return False
        return b'encryption-data' in zf.read('META-INF/manifest.xml')
```

`hwpx_edit.py`는 `is_encrypted_hwpx()`를 내부적으로 호출하고, 암호화된 파일 입력 시 해제 절차를 담은 오류 메시지를 출력한다.

## 해제 스크립트 (한컴 COM)

한컴오피스가 설치된 Windows에서만 가능.

```python
import os
import win32com.client as win32
import pythoncom

INPUT = r"<암호화된.hwpx 절대경로>"
OUTPUT = r"<해제결과.hwpx 절대경로>"
PASSWORD = "<비밀번호>"

pythoncom.CoInitialize()
try:
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.XHwpWindows.Item(0).Visible = False
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    hwp.SetMessageBoxMode(0x00020000)  # 팝업 숨김

    # 1) 비밀번호를 Open()의 옵션 문자열로 전달
    hwp.Open(os.path.abspath(INPUT), "HWPX", f"password:{PASSWORD}")

    # 2) FilePasswordChange 액션으로 암호 제거
    act = hwp.CreateAction("FilePasswordChange")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("String", "")          # 새 암호: 빈 문자열 = 제거
    pset.SetItem("Ask", 0)              # 0 = 대화상자 없이 적용
    pset.SetItem("ReadString", "")      # 열기 암호
    pset.SetItem("WriteString", "")     # 쓰기 암호
    pset.SetItem("RWAsk", 0)
    act.Execute(pset)

    # 3) 평문 HWPX로 저장
    hwp.SaveAs(os.path.abspath(OUTPUT), "HWPX", "")
    hwp.Quit()
finally:
    pythoncom.CoUninitialize()
```

## 검증

```bash
# 해제 후 manifest.xml에 encryption-data가 사라졌는지 확인
unzip -p output.hwpx META-INF/manifest.xml | grep -c encryption-data
# 0 이 나와야 정상. 양수면 암호가 여전히 남아있음.
```

## 함정

1. **단순 SaveAs만으로는 암호가 제거되지 않는다.** 문서를 비밀번호로 열어도 `SaveAs('HWPX', '')`는 원래 암호를 그대로 재적용한다. **반드시 `FilePasswordChange` 액션을 먼저 실행해야** 평문으로 저장된다.

2. **`Open()` 세 번째 인자의 형식**: `f"password:{PASSWORD}"` 문자열 (key:value 형식). 공식 문서에는 잘 안 보이지만 실측 확인됨.

3. **ParameterSet 초기화**: `act.CreateSet()`만 하면 필드가 비어있을 수 있어 `act.GetDefault(pset)`로 기본값을 채운 뒤 덮어쓰는 순서가 안전하다.

4. **`Ask=0`이 핵심**: `Ask=1`이면 한글이 "암호 확인" 대화상자를 띄우려 하고 `-NonInteractive` 환경에서 멈춘다.

5. **보안 DLL 경로 검증**: 레지스트리에 등록된 `FilePathCheckerModuleExample.dll` 경로에 실제 파일이 있어야 한다. 경로는 있지만 파일이 없으면 `RegisterModule`이 조용히 실패하고 보안 팝업이 매 Open/SaveAs마다 뜬다.
   ```bash
   reg query "HKCU\SOFTWARE\HNC\HwpAutomation\Modules"  # 등록 경로 확인
   ls -la "<등록경로>"                                   # 파일 존재 확인
   ```

6. **HWPX 전용**: `FilePasswordChange`는 HWPX뿐 아니라 HWP에도 동작하지만, 이 스킬의 주 대상은 HWPX다. HWP면 먼저 `hwp2hwpx.bat`로 변환해도 되지만 변환 자체가 암호 해제 효과도 있는지는 확인되지 않았다.

7. **COM 프로세스 잔류**: 해제 도중 에러 발생 시 `Hwp.exe` 프로세스가 남을 수 있다. 다음 실행 전 `taskkill /F /IM Hwp.exe`로 정리.

## 참고 ParameterSet — `Password` (ID 92)

| 필드 | 타입 | 의미 |
|------|------|------|
| `String` | BSTR | 새 암호(설정용) 또는 확인용 암호(Ask=TRUE) |
| `Ask` | UI1 | TRUE = 암호 확인, FALSE = 암호 설정/변경 |
| `RWAsk` | UI1 | 읽기/쓰기 암호 모드 플래그 |
| `ReadString` | BSTR | 열기 암호 |
| `WriteString` | BSTR | 쓰기 암호 |

> 출처: `reference/parameterset-table.md` (92) Password

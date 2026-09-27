# 한컴 COM 자동화 (Windows 전용, 한컴오피스 필수)

한컴오피스가 설치된 Windows 환경에서 `HWPFrame.HwpObject` COM을 통해 HWPX를 조작할 수 있다.
XML 직접 편집으로 불가능한 작업(이미지 삽입, PDF 변환)과 COM 네이티브 문서 생성에 사용. 함정 목록은 `reference/warnings-com.md`.

## 목차

- [두 파이프라인 분리 원칙](#두-파이프라인-분리-원칙)
- [hwpx_com.py 사용법 (pyhwpx 기반)](#hwpx_compy-사용법-pyhwpx-기반)
- [빠른 진단 및 PDF 변환](#빠른-진단-및-pdf-변환)
- [원격 실행 (SSH)](#원격-실행-ssh)
- [보안모듈 (팝업 제거)](#보안모듈-팝업-제거)
- [암호화된 HWPX 해제 (비밀번호 필요)](#암호화된-hwpx-해제-비밀번호-필요)
- [기본 패턴](#기본-패턴)
- [이미지 삽입](#이미지-삽입)
- [이미지 위치/배치 변경: XML 후처리 방식 (권장)](#이미지-위치배치-변경-xml-후처리-방식-권장)
- [텍스트 검색](#텍스트-검색)
- [주요 주의사항](#주요-주의사항)


## 두 파이프라인 분리 원칙

이 스킬은 두 백엔드를 가진다:

| 파이프라인 | 도구 | 환경 | 적합 작업 |
|-----------|------|------|----------|
| XML/Pandoc | hwpx_edit.py, hwpx_convert.py, python-hwpx | 크로스플랫폼, 한컴 불필요 | 기존 문서 편집, 대량 처리, 서버 |
| COM 네이티브 | **hwpx_com.py** (pyhwpx) | Windows + 한컴오피스 | 신규 생성+COM 후속 작업, 이미지 삽입, PDF |

- **COM 생성 HWPX → XML 파이프라인 읽기·편집**: 안전. COM으로 넣은 결과를 XML로 후처리하고 다시 COM으로 여는 흐름(아래 "이미지 위치/배치 변경", `hwpx_sign.py`)도 이 범위다. → 사례
- **Pandoc 생성 HWPX → COM**: 금지. COM이 본문을 0자로 인식하고, 같은 파일에 저장하면 내용 전체가 사라진다(`reference/warnings-com.md` 4번). `--normalize`·`--insert-image`·Windows `--to-pdf`(`hwpx_edit.py`·`hwpx_com.py` 모두)처럼 COM이 여는 명령에도 넣지 않는다. 한컴 결과물이 필요하면 `--from-md`로 COM에서 생성한다.
- hwpx_com.py는 입력 파일을 절대 in-place로 덮어쓰지 않는다 (위 사고의 구조적 방지)

## hwpx_com.py 사용법 (pyhwpx 기반)

```bash
# pyhwpx + COM 진단 (기동·종료까지 실측)
python hwpx_com.py --diagnose

# Markdown 부분집합 → COM 네이티브 HWPX 생성
# 지원: 제목 #/##/### (굵게 16/14/12pt), 본문 단락, **굵게** 인라인, 불릿(-/*), 파이프 표
python hwpx_com.py output.hwpx --from-md input.md

# 문서 끝에 이미지 삽입 (별도 파일 <입력>_img.hwpx로 저장, mm 단위 크기 지정 가능)
python hwpx_com.py doc.hwpx --insert-image stamp.png [--image-width 40 --image-height 20] [-o out.hwpx]

# COM 기준 본문 텍스트 추출: Pandoc HWPX 호환성 점검에도 사용 (0자면 비호환)
python hwpx_com.py doc.hwpx --get-text

# 한컴 재저장 정규화: python-hwpx로 생성·편집한 HWPX(Pandoc 변환물 제외)를 한컴이
# 직접 재저장해 lineseg·미리보기(PrvText/PrvImage)·내부 캐시를 네이티브로 재계산.
# 인쇄·배포 직전 마무리 표준 단계. in-place 금지. 본문 자수 보존을 저장 전후로 비교하지만,
# 입력 본문을 COM이 0자로 읽으면(Pandoc 변환물) 비교가 통과해 "0자 → 0자"로 완료를 낸다
python hwpx_com.py doc.hwpx --normalize [-o final.hwpx]

# PDF 저장 (hwpx_edit.py --to-pdf와 동일 기능, COM 작업 연속 시 이쪽 사용)
python hwpx_com.py doc.hwpx --to-pdf [-o out.pdf] [--password "암호"]
```

> **pyhwpx 함정**: `get_text_file()` 래퍼는 기본 option이 `saveblock:true`(선택 블록만)라서
> 선택이 없으면 None을 반환한다. 전체 본문은 저수준 `hwp.hwp.GetTextFile("TEXT", "")` 호출
> (hwpx_com.py `--get-text`가 이 방식). 상세: `reference/warnings-com.md` 7번

## 빠른 진단 및 PDF 변환

```bash
# pywin32, 보안 모듈 레지스트리, HWPFrame.HwpObject 생성 가능 여부 확인
python hwpx_edit.py --diagnose-com

# 기본 출력: 원본 폴더의 _work-hwpx-automation/<파일명>.pdf (원본 비파괴)
# PDF가 최종 결과물이면 작업 폴더로 옮기거나 -o로 작업 폴더를 직접 지정
python hwpx_edit.py <파일.hwpx> --to-pdf

# 출력 경로 지정
python hwpx_edit.py <파일.hwpx> --to-pdf -o <파일.pdf>

# 암호화된 HWPX/HWP를 PDF로 저장
python hwpx_edit.py <파일.hwpx> --to-pdf --password "<비밀번호>" -o <파일.pdf>
```

## 원격 실행 (SSH)

⚠ **SSH 세션에서는 COM 서버가 뜨지 않는다**(`서버 실행이 실패했습니다`, 대화형 데스크톱 없음). 원격에서 돌릴 때는 `schtasks /Create /TN <이름> /SC ONCE /ST 00:00 /IT /RL HIGHEST /TR "cmd /c cd /d <작업폴더> && python hwpx_com.py ... > com.log 2>&1" /F` 후 `schtasks /Run`으로 로그인 세션에서 실행하고 결과 파일을 폴링한다. `hwpx_sign.py`·`hwpx_edit.py --to-pdf`도 같다. → 사례

## 보안모듈 (팝업 제거)

한컴 COM으로 파일을 열거나 저장할 때 "접근 허용" 보안 대화상자가 반복 표시된다. 이를 제거하려면:

1. **DLL**: `<repo>/vendor/FilePathCheckerModuleExample.dll`(저장소에 커밋되지 않음, 한컴 독점 라이선스). 설치: `vendor/README.md` 참조
2. **레지스트리**: `HKCU\SOFTWARE\HNC\HwpAutomation\Modules` → 문자열 값 `FilePathCheckerModule` = DLL 전체 경로
3. **코드**: `hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')`

> 다운로드 출처: `https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/보안모듈(Automation).zip` → 압축 해제 후 `vendor/`에 배치 → 레지스트리 등록

레지스트리 경로와 실제 DLL 위치가 어긋나면 팝업이 매 Open/SaveAs마다 뜬다. 진단: `reference/warnings-com.md` 5번.

## 암호화된 HWPX 해제 (비밀번호 필요)

`Open()`의 세 번째 인자(`password:<암호>`)로 연 뒤 `FilePasswordChange` 액션으로 암호를 제거해야 한다. 단순 `SaveAs`만 하면 암호화된 채로 저장된다. 감지·해제 스크립트 전체: `reference/encrypted-hwpx.md`.

## 기본 패턴

```python
import win32com.client as win32
import pythoncom, os, time

pythoncom.CoInitialize()
hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.XHwpWindows.Item(0).Visible = False
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')

hwp.Open(os.path.abspath('input.hwpx'), 'HWPX', '')
# ... 작업 ...
hwp.SaveAs(os.path.abspath('output.hwpx'), 'HWPX', '')
hwp.SaveAs(os.path.abspath('output.pdf'), 'PDF', '')
hwp.Quit()
pythoncom.CoUninitialize()
```

## 이미지 삽입

```python
# 커서를 원하는 위치로 이동
hwp.HAction.Run('MoveDocEnd')

# 이미지 삽입 (Width/Height 단위: mm)
ctrl = hwp.InsertPicture(abs_img_path, True, 1, False, False, 0, 23, 15)
# params: path, Embedded, sizeoption(1=지정크기), Reverse, Watermark, Effect, Width_mm, Height_mm
```

## 이미지 위치/배치 변경: XML 후처리 방식 (권장)

COM의 `ShapeObjDialog`로 속성 변경 시 `TreatAsChar`, `TextWrap` 등이 정상 반영되지 않는 경우가 많다.
**권장 워크플로우**: COM으로 이미지 삽입 → HWPX 저장 → XML 후처리로 `<hp:pic>` 속성 변경 → 다시 COM으로 PDF 저장.

```python
# XML에서 hp:pic의 pos/sz/textWrap 속성을 직접 수정
# 수동 완성본의 <hp:pic> 요소를 분석하여 좌표/크기를 복사하는 것이 가장 정확
import re
section_xml = re.sub(
    r'(<hp:pic[^>]*?)textWrap="[^"]*"',
    r'\1textWrap="IN_FRONT_OF_TEXT"', section_xml)
section_xml = re.sub(
    r'(<hp:pic[^>]*>.*?)<hp:pos [^/]*/>(.*?</hp:pic>)',
    r'\1<hp:pos treatAsChar="0" ... vertRelTo="PAPER" horzRelTo="PAPER" '
    r'vertOffset="NNN" horzOffset="NNN"/>\2',
    section_xml, flags=re.DOTALL)
```

## 텍스트 검색

```python
fr = hwp.HParameterSet.HFindReplace
hwp.HAction.GetDefault('RepeatFind', fr.HSet)
fr.FindString = '검색어'
fr.MatchCase = 1
fr.ReplaceMode = 0  # find only
hwp.HAction.Execute('RepeatFind', fr.HSet)
# 이후 커서 이동: hwp.HAction.Run('MoveRight') 등
```

## 주요 주의사항

1. **COM 프로세스 잔류**: 에러 발생 시 `hwp.Quit()`이 호출되지 않아 `Hwp.exe`가 잔류한다. 정리 절차(사용자가 연 문서 확인 뒤 강제 종료)는 `reference/warnings-com.md` 9번
2. **COM 재초기화 금지**: 동일 프로세스에서 `CoUninitialize()` 후 `CoInitialize()`를 다시 호출하면 segfault 발생 가능. COM 작업을 2단계로 나눠야 하면 **별도 Python 스크립트**로 분리
3. **HParameterSet 속성명**: COM 타입 라이브러리에 따라 속성명이 다름. `dir(obj)` 또는 `_prop_map_put_`으로 확인. 예: `IgnoreCase` → `MatchCase`, `FileName` → `filename`
4. **InsertPicture 반환값**: 성공 시 `IDHwpCtrlCode` COM 객체, 실패 시 `None`/`False`
5. **절대 경로 필수**: COM API에 전달하는 파일 경로는 반드시 `os.path.abspath()` 사용
6. **`Close` 메서드 없음**: `hwp.Close()`는 존재하지 않음. 문서를 닫으려면 `hwp.Quit()` (앱 종료)

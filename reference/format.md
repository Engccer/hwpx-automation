# HWPX 파일 구조

HWPX는 ZIP 파일이며 내부 구조:

```
document.hwpx (ZIP)
├── mimetype                    # "application/hwp+zip"
├── META-INF/
│   ├── container.xml           # 루트 파일 지정
│   ├── container.rdf           # 메타데이터
│   └── manifest.xml            # 파일 목록
├── Contents/
│   ├── content.hpf             # 콘텐츠 매니페스트
│   ├── header.xml              # 문서 설정 (글꼴, 스타일, borderFill 등)
│   └── section0.xml            # 본문 내용 ★ 편집 대상
├── Preview/
│   ├── PrvImage.png            # 미리보기 이미지
│   └── PrvText.txt             # 미리보기 텍스트
├── settings.xml                # 편집 설정
└── version.xml                 # 버전 정보
```

## section0.xml 주요 요소

```xml
<hs:sec>                        <!-- 섹션 -->
  <hp:p>                         <!-- 단락 (paragraph) -->
    <hp:run>                     <!-- 텍스트 런 -->
      <hp:rPr>                   <!-- 런 속성 (글꼴, 크기 등) -->
      <hp:t>텍스트 내용</hp:t>    <!-- 실제 텍스트 ★ -->
    </hp:run>
  </hp:p>
  <hp:tbl>                       <!-- 표 -->
    <hp:tr>                      <!-- 행 -->
      <hp:tc>                    <!-- 셀 -->
        <hp:cellAddr             <!-- 셀 주소 -->
          rowIndex="0" colIndex="0"
          rowSpan="1" colSpan="1"/>
        <hp:cellSz width="..." height="..."/>
        <hp:p>                   <!-- 셀 안의 단락 -->
          <hp:run>
            <hp:t>셀 내용</hp:t>
          </hp:run>
        </hp:p>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>
```

## 네임스페이스

```python
NS = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
}
```

- HWPML 2016 → 2011 네임스페이스 자동 변환 지원 (python-hwpx v2.8+)
- 다중 섹션 문서: `section1.xml`, `section2.xml` 등 존재 가능

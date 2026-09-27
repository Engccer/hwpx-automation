# HWPX 파일 구조

HWPX는 ZIP 파일이며 내부 구조:

```
document.hwpx (ZIP)
├── mimetype                    # "application/hwp+zip" (첫 항목, ZIP_STORED)
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
<hs:sec>                                  <!-- 섹션 -->
  <hp:p paraPrIDRef="0" styleIDRef="0">   <!-- 단락 (paragraph) -->
    <hp:run charPrIDRef="0">              <!-- 텍스트 런 (글자 속성은 header.xml charPr id 참조) -->
      <hp:t>텍스트 내용</hp:t>             <!-- 실제 텍스트 ★ -->
    </hp:run>
  </hp:p>
  <hp:p>
    <hp:run>
      <hp:tbl rowCnt="2" colCnt="2">      <!-- 표: 문단의 run 안에 있다 -->
        <hp:tr>                           <!-- 행 -->
          <hp:tc borderFillIDRef="3">     <!-- 셀 -->
            <hp:subList>                  <!-- 셀 안 문단 목록 -->
              <hp:p>                      <!-- 셀 안의 단락 (최소 1개 필수) -->
                <hp:run charPrIDRef="0"><hp:t>셀 내용</hp:t></hp:run>
              </hp:p>
            </hp:subList>
            <hp:cellAddr colAddr="0" rowAddr="0"/>  <!-- 셀 좌표 -->
            <hp:cellSpan colSpan="1" rowSpan="1"/>  <!-- 병합 범위 -->
            <hp:cellSz width="..." height="..."/>
          </hp:tc>
        </hp:tr>
      </hp:tbl>
    </hp:run>
  </hp:p>
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

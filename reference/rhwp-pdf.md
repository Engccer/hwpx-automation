# macOS·Linux에서 PDF 변환 (rhwp)

한컴 COM 자동화는 Windows 전용이다. Windows가 아닌 OS에서 `hwpx_edit.py --to-pdf`는 오픈소스 조판 엔진 [rhwp](https://github.com/edwardkim/rhwp)의 `export-pdf`를 호출한다(`rhwp_pdf.py`). 한컴 없이 한글 조판 규칙으로 쪽·줄을 나눈다.

**설치**: Releases에서 플랫폼 바이너리(예: `rhwp-v0.8.6-macos-aarch64.tar.gz`)와 `SHA256SUMS.txt`를 받아 체크섬을 확인한 뒤 PATH(예: `~/.local/bin`)에 둔다. `--check-env`의 Tier 4에 `[O] rhwp`가 뜨면 준비된 것이다.

```bash
python hwpx_edit.py <파일.hwpx> --to-pdf -o <파일.pdf>
```

**충실도**: 한컴 PDF와 쪽수·글자는 같게 나오고, 남은 차이는 세 가지다. → 사례
- 글꼴이 이 컴퓨터 글꼴로 대체된다(맥: 맑은 고딕·HY헤드라인M → Apple SD Gothic Neo, 명조 → 나눔명조).
- 한 문단에서 줄 끝 단어 하나가 다음 줄로 넘어가는 수준의 줄바꿈 차이가 가끔 생긴다.
- 제목 옆 짧은 라벨의 세로 위치가 조금 다르다.

**출력 해석**:
- 표준출력 `사용 글꼴:` 목록으로 대체 여부를 확인한다.
- `경고: 쪽 하단 넘침 N건`: 조판이 쪽 끝을 몇 px 넘었다는 뜻이다. 쪽수는 대개 그대로지만 쪽 끝 줄을 렌더로 확인한다.
- **exit 2 + `글꼴에 없는 문자가 빈 네모로 찍혔습니다: ✅(U+2705)`**: PDF는 남지만 배포하면 안 된다. 이모지·특수기호가 설치 글꼴에 없어 LastResort(빈 네모)로 찍힌 것이다. v0.8.6에서는 `--font-path`·`RHWP_FONT_PATH`·Noto Emoji 사용자 설치 모두 이 대체를 바꾸지 못했다. 해당 문자를 바꾸거나 Windows 한컴으로 변환한다.

**rhwp로 안 되는 것**: 한컴 재저장 정규화(`hwpx_com.py --normalize`), 이미지 삽입, 서명(`hwpx_sign.py`). 한컴과 똑같은 PDF가 꼭 필요하면 Windows 한컴 COM을 쓴다. 원격 Windows에 SSH로 붙을 때는 `reference/com.md` "원격 실행"을 따른다.

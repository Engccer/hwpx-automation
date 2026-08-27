# hwp2hwpx JAR 재빌드 절차

번들 `convert/hwp2hwpx-<version>-<commit>p<N>.jar`는 상류 소스에 이 폴더의 패치를 적용해 직접 빌드한다. Maven·Gradle 없이 `javac`만으로 된다(JDK 11 이상. 실측 JDK 25).

```bash
# 1. 상류 소스
git clone --depth 1 https://github.com/neolord0/hwp2hwpx.git
cd hwp2hwpx && git log -1 --format=%h      # 파일명에 넣을 커밋 해시

# 2. 패치 적용 (상류에 이미 반영됐으면 건너뜀)
git apply ../0001-extendControl-range-guard.patch

# 3. 의존 JAR (pom.xml의 hwplib·hwpxlib 버전과 맞춘다)
curl -sSLO https://repo1.maven.org/maven2/kr/dogfoot/hwpxlib/1.0.9/hwpxlib-1.0.9.jar
# hwplib은 convert/lib/hwplib-1.1.10.jar 재사용

# 4. 컴파일·패키징
mkdir -p build/classes
find src/main/java -name "*.java" > build/sources.txt
javac --release 11 -encoding UTF-8 -nowarn \
  -cp "../lib/hwplib-1.1.10.jar:hwpxlib-1.0.9.jar" -d build/classes @build/sources.txt
(cd build/classes && jar cf ../hwp2hwpx-1.0.0-<commit>p1.jar kr)

# 5. 교체: convert/ 의 옛 JAR 삭제, hwp2hwpx.sh·hwp2hwpx.bat 의 CP/JAR1·JAR3 파일명 갱신
```

검증: 패치 계기가 된 문서(머리글에 컨트롤 불일치가 있는 554쪽 연구보고서, 외부 자료라 저장소에 두지 않음)를 변환해 `hwpx_edit.py --to-md` 결과가 한컴 COM 변환본과 글자·표 행 단위로 일치하는지 확인한다(2026-08-27 기준 누락 0·잉여 0).

## 패치 목록

| 파일 | 내용 | 상류 반영 |
|---|---|---|
| `0001-extendControl-range-guard.patch` | `ForChars.extendControl`에서 컨트롤 인덱스가 컨트롤 목록 범위를 넘으면 건너뜀(머리글·바닥글 문단의 컨트롤 문자·데이터 수 불일치) | 미반영(상류 #5는 null 체크만) |

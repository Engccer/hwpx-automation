#!/bin/sh
# HWP -> HWPX 변환 (macOS/Linux용. Windows는 같은 폴더의 hwp2hwpx.bat 사용)
# 사용법: hwp2hwpx.sh <input.hwp> [output.hwpx]
# 출력 미지정 시 입력 폴더의 _work-hwpx-automation/ 하위에 <이름>.hwpx로 생성(원본 비파괴).
set -u

TOOL_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)

# --- JDK 탐색: JAVA_HOME > PATH의 java ---
JAVA_BIN=""
if [ -n "${JAVA_HOME:-}" ] && [ -x "$JAVA_HOME/bin/java" ]; then
    JAVA_BIN="$JAVA_HOME/bin/java"
elif command -v java >/dev/null 2>&1; then
    JAVA_BIN="java"
else
    echo "ERROR: JDK 21 not found. Install JDK 21 or set JAVA_HOME." >&2
    exit 1
fi

INPUT="${1:-}"
OUTPUT="${2:-}"

if [ -z "$INPUT" ]; then
    echo "Usage: $(basename -- "$0") <input.hwp> [output.hwpx]" >&2
    exit 2
fi

if [ ! -f "$INPUT" ]; then
    echo "ERROR: input file not found: $INPUT" >&2
    exit 1
fi

if [ -z "$OUTPUT" ]; then
    INPUT_DIR=$(CDPATH= cd -- "$(dirname -- "$INPUT")" && pwd)
    BASE=$(basename -- "$INPUT")
    NAME="${BASE%.*}"
    OUTPUT_DIR="$INPUT_DIR/_work-hwpx-automation"
    mkdir -p "$OUTPUT_DIR"
    OUTPUT="$OUTPUT_DIR/$NAME.hwpx"
fi

# macOS/Linux JVM은 UTF-8 로케일에서 argv를 그대로 받으므로
# .bat의 임시 폴더 staging(cp949 우회)이 여기서는 필요 없다.
CP="$TOOL_DIR/hwp2hwpx-1.0.0.jar:$TOOL_DIR/lib/hwplib-1.1.10.jar:$TOOL_DIR/lib/hwpxlib-1.0.8.jar:$TOOL_DIR"

exec "$JAVA_BIN" -cp "$CP" Hwp2HwpxCLI "$INPUT" "$OUTPUT"

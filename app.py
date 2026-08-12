from flask import Flask, render_template, request, jsonify
import sqlite3
from datetime import datetime
import os
import re


app = Flask(__name__)


# =========================================================
# DB 설정
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "barcode.db")


# =========================================================
# DB 초기화 / 기존 DB 자동 보완
# =========================================================

def init_db():

    conn = sqlite3.connect(DB_NAME)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS scans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            main_category TEXT,
            sub_category TEXT,
            alc_code TEXT,
            alc_type TEXT,
            p_code TEXT,
            part_number TEXT,
            product_date TEXT,
            scanned_at TEXT
        )
    """)

    # 기존 DB를 사용하고 있는 경우
    # 없는 컬럼만 자동으로 추가
    columns = [
        row[1]
        for row in conn.execute(
            "PRAGMA table_info(scans)"
        ).fetchall()
    ]

    required_columns = {
        "main_category": "TEXT",
        "sub_category": "TEXT",
        "alc_code": "TEXT",
        "alc_type": "TEXT",
        "p_code": "TEXT",
        "part_number": "TEXT",
        "product_date": "TEXT",
        "scanned_at": "TEXT"
    }

    for column_name, column_type in required_columns.items():

        if column_name not in columns:

            conn.execute(
                f"""
                ALTER TABLE scans
                ADD COLUMN {column_name} {column_type}
                """
            )

    conn.commit()
    conn.close()


# =========================================================
# 바코드 분석
# =========================================================

def parse_barcode(raw):

    alc_code = ""
    alc_type = ""

    p_code = ""
    part_number = ""
    product_date = ""

    if not raw:

        return (
            alc_code,
            alc_type,
            p_code,
            part_number,
            product_date
        )


    # ---------------------------------------------------------
    # Data Matrix 안에 여러 레코드가 #으로 붙어 있는 경우
    # 각각 분리
    # ---------------------------------------------------------

    records = raw.split("#")


    for record in records:

        # ASCII GS = \x1d
        fields = record.split("\x1d")


        # 이 레코드에서 찾은 값
        record_alc = ""
        record_alc_type = ""
        record_p = ""
        record_part = ""
        record_date = ""


        for field in fields:

            field = field.strip()

            if not field:
                continue


            # =================================================
            # ALC 찾기
            # =================================================
            #
            # 중요:
            #
            # ALC는 반드시 S 필드에서만 찾음
            #
            # SU304 -> U304 -> FRT
            # SL123 -> L123 -> FRT
            #
            # SR148 -> R148 -> RR
            # SRA8  -> RA8  -> RR
            # SS123 -> S123 -> RR
            #
            # VR148 같은 V 필드는 절대 ALC로 사용하지 않음
            # =================================================

            if (
                not record_alc
                and field.startswith("S")
                and len(field) >= 2
            ):

                candidate = (
                    field[1:]
                    .strip()
                    .upper()
                )


                # U/L/R/S 로 시작하는
                # 영문+숫자 조합만 인정
                if re.fullmatch(
                    r"[ULRS][A-Z0-9]+",
                    candidate
                ):

                    record_alc = candidate


                    if candidate[0] in ("U", "L"):

                        record_alc_type = "FRT"


                    elif candidate[0] in ("R", "S"):

                        record_alc_type = "RR"


            # =================================================
            # P 코드
            # =================================================
            #
            # P82302XB000NNB
            # ->
            # 82302XB000NNB
            # =================================================

            if (
                not record_p
                and field.startswith("P")
                and len(field) > 1
            ):

                candidate = field[1:].strip()

                if candidate:

                    record_p = candidate


            # =================================================
            # 부품번호
            # =================================================
            #
            # CL4-EEC87-00000107
            # ->
            # L4-EEC87-00000107
            # =================================================

            if (
                not record_part
                and field.startswith("CL4-")
            ):

                record_part = (
                    field[1:]
                    .strip()
                )


            # =================================================
            # 생산일자
            # =================================================
            #
            # T2608074010A00000923
            #
            # 첫 6자리:
            # 260807
            #
            # ->
            # 2026-08-07
            # =================================================

            if (
                not record_date
                and field.startswith("T")
                and len(field) >= 7
            ):

                date_text = field[1:7]


                if date_text.isdigit():

                    year = "20" + date_text[0:2]
                    month = date_text[2:4]
                    day = date_text[4:6]


                    try:

                        parsed_date = datetime.strptime(
                            f"{year}-{month}-{day}",
                            "%Y-%m-%d"
                        )

                        record_date = (
                            parsed_date.strftime(
                                "%Y-%m-%d"
                            )
                        )

                    except ValueError:

                        pass


        # =====================================================
        # 정상적인 ALC가 있는 레코드만 선택
        # =====================================================
        #
        # 이렇게 해야 뒤에 붙은 다른 Data Matrix 정보가
        # 현재 제품의 ALC 수량에 같이 들어가는 것을 막을 수 있음.
        # =====================================================

        if record_alc:

            alc_code = record_alc
            alc_type = record_alc_type

            p_code = record_p
            part_number = record_part
            product_date = record_date

            break


    # =========================================================
    # 부품번호 보조 검색
    # =========================================================

    if not part_number:

        match = re.search(
            r"L4-[A-Za-z0-9]+-\d+",
            raw
        )

        if match:

            part_number = match.group(0)


    # =========================================================
    # 일반 날짜 형식 보조 검색
    # =========================================================

    if not product_date:

        match = re.search(
            r"(20\d{2})[-./](\d{1,2})[-./](\d{1,2})",
            raw
        )

        if match:

            year = match.group(1)
            month = match.group(2).zfill(2)
            day = match.group(3).zfill(2)

            try:

                parsed_date = datetime.strptime(
                    f"{year}-{month}-{day}",
                    "%Y-%m-%d"
                )

                product_date = (
                    parsed_date.strftime(
                        "%Y-%m-%d"
                    )
                )

            except ValueError:

                pass


    return (
        alc_code,
        alc_type,
        p_code,
        part_number,
        product_date
    )


# =========================================================
# 통계 가져오기
# =========================================================

def get_statistics(conn):

    conn.row_factory = sqlite3.Row

    today = datetime.now().strftime("%Y-%m-%d")


    # =====================================================
    # 당일생산 완제품
    # =====================================================

    today_finished = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '당일생산분'
        AND sub_category = '완제품'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    # =====================================================
    # 당일생산 반제품
    # =====================================================

    today_semi = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '당일생산분'
        AND sub_category = '반제품'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    # =====================================================
    # 당일 생산량
    # =====================================================

    today_production = (
        today_finished +
        today_semi
    )


    # =====================================================
    # 재고 완제품
    # =====================================================

    stock_finished = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '재고'
        AND sub_category = '완제품'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    # =====================================================
    # 재고 반제품
    # =====================================================

    stock_semi = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '재고'
        AND sub_category = '반제품'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    # =====================================================
    # 오늘 출고
    # =====================================================

    shipped = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '출고'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    # =====================================================
    # 오늘 불량
    # =====================================================

    defective = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '불량'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    # =====================================================
    # FRT 전체
    # =====================================================

    frt_total = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE alc_type = 'FRT'
    """).fetchone()[0]


    # =====================================================
    # RR 전체
    # =====================================================

    rr_total = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE alc_type = 'RR'
    """).fetchone()[0]


    # =====================================================
    # FRT 오늘
    # =====================================================

    frt_today = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE alc_type = 'FRT'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    # =====================================================
    # RR 오늘
    # =====================================================

    rr_today = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE alc_type = 'RR'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    # =====================================================
    # ALC 전체 수량
    # =====================================================

    alc_counts_rows = conn.execute("""
        SELECT
            alc_type,
            alc_code,
            COUNT(*) AS count
        FROM scans
        WHERE alc_code IS NOT NULL
        AND alc_code != ''
        GROUP BY alc_type, alc_code
        ORDER BY alc_type, alc_code
    """).fetchall()


    # =====================================================
    # 오늘 ALC 수량
    # =====================================================

    today_alc_rows = conn.execute("""
        SELECT
            alc_type,
            alc_code,
            COUNT(*) AS count
        FROM scans
        WHERE alc_code IS NOT NULL
        AND alc_code != ''
        AND DATE(scanned_at) = ?
        GROUP BY alc_type, alc_code
        ORDER BY alc_type, alc_code
    """, (today,)).fetchall()


    # =====================================================
    # 생산일자별 수량
    # =====================================================

    date_rows = conn.execute("""
        SELECT
            product_date,
            COUNT(*) AS count
        FROM scans
        WHERE product_date IS NOT NULL
        AND product_date != ''
        GROUP BY product_date
        ORDER BY product_date DESC
    """).fetchall()


    # =====================================================
    # 최근 스캔 20개
    # =====================================================

    scan_rows = conn.execute("""
        SELECT *
        FROM scans
        ORDER BY id DESC
        LIMIT 20
    """).fetchall()


    # =====================================================
    # JSON으로 보내기 쉽게 변환
    # =====================================================

    alc_counts = [
        dict(row)
        for row in alc_counts_rows
    ]

    today_alc_counts = [
        dict(row)
        for row in today_alc_rows
    ]

    date_counts = [
        dict(row)
        for row in date_rows
    ]

    scans = [
        dict(row)
        for row in scan_rows
    ]


    return {
        "today_production": today_production,

        "today_finished": today_finished,
        "today_semi": today_semi,

        "stock_finished": stock_finished,
        "stock_semi": stock_semi,

        "shipped": shipped,
        "defective": defective,

        "frt_total": frt_total,
        "rr_total": rr_total,

        "frt_today": frt_today,
        "rr_today": rr_today,

        "alc_counts": alc_counts,
        "today_alc_counts": today_alc_counts,
        "date_counts": date_counts,
        "scans": scans
    }


# =========================================================
# 메인 화면
# =========================================================

@app.route("/")
def index():

    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row

    stats = get_statistics(conn)

    conn.close()


    return render_template(
        "index.html",

        today_production=
            stats["today_production"],

        today_finished=
            stats["today_finished"],

        today_semi=
            stats["today_semi"],

        stock_finished=
            stats["stock_finished"],

        stock_semi=
            stats["stock_semi"],

        shipped=
            stats["shipped"],

        defective=
            stats["defective"],

        frt_total=
            stats["frt_total"],

        rr_total=
            stats["rr_total"],

        frt_today=
            stats["frt_today"],

        rr_today=
            stats["rr_today"],

        alc_counts=
            stats["alc_counts"],

        today_alc_counts=
            stats["today_alc_counts"],

        date_counts=
            stats["date_counts"],

        scans=
            stats["scans"]
    )


# =========================================================
# 바코드 스캔
# =========================================================

@app.route("/scan", methods=["POST"])
def scan():

    try:

        data = request.get_json()


        if not data:

            return jsonify({
                "success": False,
                "error": "전송된 데이터가 없습니다."
            }), 400


        # =================================================
        # 데이터
        # =================================================

        raw = data.get(
            "raw",
            ""
        )

        main_category = data.get(
            "main_category",
            ""
        )

        sub_category = data.get(
            "sub_category",
            ""
        )


        # =================================================
        # 카테고리 확인
        # =================================================

        allowed_main = [
            "당일생산분",
            "재고",
            "출고",
            "불량"
        ]


        if main_category not in allowed_main:

            return jsonify({
                "success": False,
                "error": "올바른 카테고리를 선택해주세요."
            }), 400


        if main_category in (
            "당일생산분",
            "재고"
        ):

            if sub_category not in (
                "완제품",
                "반제품"
            ):

                return jsonify({
                    "success": False,
                    "error": "완제품 또는 반제품을 선택해주세요."
                }), 400


        else:

            sub_category = ""


        if not raw:

            return jsonify({
                "success": False,
                "error": "바코드 데이터를 읽지 못했습니다."
            }), 400


        # =================================================
        # 콘솔 원본 출력
        # =================================================

        print()
        print("==============================")
        print("바코드 원본 데이터")
        print("==============================")
        print(repr(raw))
        print("==============================")


        # =================================================
        # 바코드 분석
        # =================================================

        (
            alc_code,
            alc_type,
            p_code,
            part_number,
            product_date
        ) = parse_barcode(raw)


        # =================================================
        # ALC가 없으면 저장하지 않음
        # =================================================

        if not alc_code:

            print("ALC 인식 실패")

            return jsonify({
                "success": False,
                "error":
                    "ALC 코드를 찾지 못했습니다.\n"
                    "S 필드의 U/L/R/S 코드를 확인해주세요."
            }), 400


        # =================================================
        # 스캔 시간
        # =================================================

        scanned_at = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        # =================================================
        # DB 저장
        # =================================================

        conn = sqlite3.connect(DB_NAME)
        conn.row_factory = sqlite3.Row


        conn.execute("""
            INSERT INTO scans (
                main_category,
                sub_category,
                alc_code,
                alc_type,
                p_code,
                part_number,
                product_date,
                scanned_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            main_category,
            sub_category,
            alc_code,
            alc_type,
            p_code,
            part_number,
            product_date,
            scanned_at
        ))


        conn.commit()


        # =================================================
        # 현재 ALC 누적 수량
        # =================================================

        alc_total = conn.execute("""
            SELECT COUNT(*)
            FROM scans
            WHERE alc_code = ?
        """, (
            alc_code,
        )).fetchone()[0]


        # =================================================
        # 현재 ALC 오늘 수량
        # =================================================

        today = datetime.now().strftime(
            "%Y-%m-%d"
        )


        alc_today = conn.execute("""
            SELECT COUNT(*)
            FROM scans
            WHERE alc_code = ?
            AND DATE(scanned_at) = ?
        """, (
            alc_code,
            today
        )).fetchone()[0]


        # =================================================
        # 전체 통계
        # =================================================

        stats = get_statistics(conn)

        conn.close()


        # =================================================
        # 콘솔 출력
        # =================================================

        print("==============================")
        print("저장 완료")
        print("업무구분:", main_category)
        print("제품종류:", sub_category or "-")
        print("ALC:", alc_code)
        print("ALC 구분:", alc_type)
        print("P코드:", p_code or "-")
        print("부품번호:", part_number or "-")
        print("생산일자:", product_date or "-")
        print("스캔시간:", scanned_at)
        print("ALC 누적:", alc_total)
        print("ALC 오늘:", alc_today)
        print("==============================")
        print()


        # =================================================
        # 브라우저 응답
        # =================================================

        response_data = {

            "success": True,

            "main_category":
                main_category,

            "sub_category":
                sub_category,

            "alc_code":
                alc_code,

            "alc_type":
                alc_type,

            "alc_total":
                alc_total,

            "alc_today":
                alc_today,

            "p_code":
                p_code,

            "part_number":
                part_number,

            "product_date":
                product_date,

            "scanned_at":
                scanned_at
        }


        response_data.update(stats)


        return jsonify(
            response_data
        )


    except Exception as e:

        print()
        print("==============================")
        print("오류 발생")
        print("==============================")
        print(str(e))
        print("==============================")
        print()


        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# =========================================================
# 실행
# =========================================================

# =========================================================
# DB 초기화
# =========================================================

init_db()


# =========================================================
# 로컬 실행
# =========================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
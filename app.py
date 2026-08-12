from flask import Flask, render_template, request, jsonify
import sqlite3
from datetime import datetime
import os
import re


app = Flask(__name__)


# =========================================================
# DB
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "barcode.db")


# =========================================================
# DB 초기화
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

    # 여러 Data Matrix가 #으로 붙은 경우
    records = raw.split("#")


    for record in records:

        fields = record.split("\x1d")

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
            # ALC
            #
            # 반드시 S 필드만 사용
            #
            # SU304 -> U304 -> FRT
            # SL123 -> L123 -> FRT
            #
            # SRA8  -> RA8  -> RR
            # SR148 -> R148 -> RR
            # SS123 -> S123 -> RR
            #
            # VR148 등 V 필드는 무시
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

            if (
                not record_p
                and field.startswith("P")
                and len(field) > 1
            ):

                record_p = field[1:].strip()


            # =================================================
            # 부품번호
            # =================================================

            if (
                not record_part
                and field.startswith("CL4-")
            ):

                record_part = field[1:].strip()


            # =================================================
            # 생산일자
            #
            # T260807...
            # -> 2026-08-07
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

                        record_date = parsed_date.strftime(
                            "%Y-%m-%d"
                        )

                    except ValueError:

                        pass


        # 정상 ALC가 있는 첫 레코드만 사용
        if record_alc:

            alc_code = record_alc
            alc_type = record_alc_type

            p_code = record_p
            part_number = record_part
            product_date = record_date

            break


    # 부품번호 보조 검색
    if not part_number:

        match = re.search(
            r"L4-[A-Za-z0-9]+-\d+",
            raw
        )

        if match:

            part_number = match.group(0)


    return (
        alc_code,
        alc_type,
        p_code,
        part_number,
        product_date
    )


# =========================================================
# 전체 통계
# =========================================================

def get_statistics(conn):

    conn.row_factory = sqlite3.Row

    today = datetime.now().strftime("%Y-%m-%d")


    # =====================================================
    # 당일생산
    # =====================================================

    today_finished = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '당일생산분'
        AND sub_category = '완제품'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    today_semi = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '당일생산분'
        AND sub_category = '반제품'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    today_production = (
        today_finished
        + today_semi
    )


    # =====================================================
    # 재고
    # =====================================================

    stock_finished = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '재고'
        AND sub_category = '완제품'
    """).fetchone()[0]


    stock_semi = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '재고'
        AND sub_category = '반제품'
    """).fetchone()[0]


    # =====================================================
    # 출고 FRT / RR - 오늘
    # =====================================================

    shipped_frt = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '출고'
        AND alc_type = 'FRT'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    shipped_rr = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '출고'
        AND alc_type = 'RR'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    shipped = (
        shipped_frt
        + shipped_rr
    )


    # =====================================================
    # 불량 FRT / RR - 오늘
    # =====================================================

    defective_frt = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '불량'
        AND alc_type = 'FRT'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    defective_rr = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE main_category = '불량'
        AND alc_type = 'RR'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    defective = (
        defective_frt
        + defective_rr
    )


    # =====================================================
    # ALC 전체 FRT/RR
    # =====================================================

    frt_total = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE alc_type = 'FRT'
    """).fetchone()[0]


    rr_total = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE alc_type = 'RR'
    """).fetchone()[0]


    frt_today = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE alc_type = 'FRT'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    rr_today = conn.execute("""
        SELECT COUNT(*)
        FROM scans
        WHERE alc_type = 'RR'
        AND DATE(scanned_at) = ?
    """, (today,)).fetchone()[0]


    # =====================================================
    # ALC별 수량
    # =====================================================

    alc_rows = conn.execute("""
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
    # 오늘 ALC
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
    # 생산일자
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
    # 최근 20개
    # =====================================================

    scan_rows = conn.execute("""
        SELECT *
        FROM scans
        ORDER BY id DESC
        LIMIT 20
    """).fetchall()


    return {

        "today_production": today_production,

        "today_finished": today_finished,
        "today_semi": today_semi,

        "stock_finished": stock_finished,
        "stock_semi": stock_semi,

        "shipped": shipped,
        "shipped_frt": shipped_frt,
        "shipped_rr": shipped_rr,

        "defective": defective,
        "defective_frt": defective_frt,
        "defective_rr": defective_rr,

        "frt_total": frt_total,
        "rr_total": rr_total,

        "frt_today": frt_today,
        "rr_today": rr_today,

        "alc_counts": [
            dict(row)
            for row in alc_rows
        ],

        "today_alc_counts": [
            dict(row)
            for row in today_alc_rows
        ],

        "date_counts": [
            dict(row)
            for row in date_rows
        ],

        "scans": [
            dict(row)
            for row in scan_rows
        ]
    }


# =========================================================
# 메인
# =========================================================

@app.route("/")
def index():

    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row

    stats = get_statistics(conn)

    conn.close()

    return render_template(
        "index.html",
        **stats
    )


# =========================================================
# 스캔
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


        raw = data.get(
            "raw",
            ""
        )


        main_category = data.get(
            "main_category",
            ""
        ).strip()


        sub_category = data.get(
            "sub_category",
            ""
        ).strip()


        # =================================================
        # 메인 카테고리
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


        # =================================================
        # 당일생산분 / 재고
        # 완제품 / 반제품
        # =================================================

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
                    "error":
                        "완제품 또는 반제품을 선택해주세요."
                }), 400


        # =================================================
        # 출고 / 불량
        # FRT / RR
        # =================================================

        elif main_category in (
            "출고",
            "불량"
        ):

            if sub_category not in (
                "FRT",
                "RR"
            ):

                return jsonify({
                    "success": False,
                    "error":
                        "FRT 또는 RR을 선택해주세요."
                }), 400


        if not raw:

            return jsonify({
                "success": False,
                "error": "바코드 데이터를 읽지 못했습니다."
            }), 400


        print()
        print("==============================")
        print("바코드 원본")
        print("==============================")
        print(repr(raw))
        print("==============================")


        # =================================================
        # 분석
        # =================================================

        (
            alc_code,
            alc_type,
            p_code,
            part_number,
            product_date
        ) = parse_barcode(raw)


        if not alc_code:

            return jsonify({
                "success": False,
                "error":
                    "ALC 코드를 찾지 못했습니다."
            }), 400


        # =================================================
        # 출고 / 불량 FRT-RR 검사
        # =================================================
        #
        # 사용자가 출고 → FRT 선택
        # 그런데 실제 바코드가 RR이면
        # 저장하지 않음.
        # =================================================

        if main_category in (
            "출고",
            "불량"
        ):

            if sub_category != alc_type:

                return jsonify({

                    "success": False,

                    "error":
                        f"ALC 구분이 일치하지 않습니다.\n\n"
                        f"선택: {sub_category}\n"
                        f"실제 바코드: {alc_type}\n"
                        f"ALC: {alc_code}"

                }), 400


        # =================================================
        # 저장
        # =================================================

        scanned_at = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


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
        # 해당 ALC 누적
        # =================================================

        alc_total = conn.execute("""
            SELECT COUNT(*)
            FROM scans
            WHERE alc_code = ?
        """, (
            alc_code,
        )).fetchone()[0]


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


        stats = get_statistics(conn)

        conn.close()


        print("==============================")
        print("저장 완료")
        print("업무:", main_category)
        print("선택:", sub_category)
        print("ALC:", alc_code)
        print("ALC 구분:", alc_type)
        print("부품번호:", part_number)
        print("생산일자:", product_date)
        print("==============================")


        response = {

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


        response.update(
            stats
        )


        return jsonify(
            response
        )


    except Exception as e:

        print(
            "오류:",
            str(e)
        )

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


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

from flask import Flask, render_template, request, jsonify, send_file, make_response
import sqlite3
from datetime import datetime
from zoneinfo import ZoneInfo
import os
import re
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR = os.environ.get("DB_DIR", BASE_DIR)
os.makedirs(DB_DIR, exist_ok=True)
DB_NAME = os.path.join(DB_DIR, "barcode.db")

KST = ZoneInfo("Asia/Seoul")


def now_kst():
    return datetime.now(KST)


def today_string():
    return now_kst().strftime("%Y-%m-%d")


def get_db():
    conn = sqlite3.connect(DB_NAME, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()

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

    columns = conn.execute("PRAGMA table_info(scans)").fetchall()
    column_names = [row["name"] for row in columns]

    required_columns = {
        "main_category": "TEXT",
        "sub_category": "TEXT",
        "alc_code": "TEXT",
        "alc_type": "TEXT",
        "p_code": "TEXT",
        "part_number": "TEXT",
        "product_date": "TEXT",
        "scanned_at": "TEXT",
    }

    for column_name, column_type in required_columns.items():
        if column_name not in column_names:
            conn.execute(
                f"ALTER TABLE scans ADD COLUMN {column_name} {column_type}"
            )

    # 과거 '재고'라는 이름을 사용한 기록을 '총재고'로 통일
    conn.execute("""
        UPDATE scans
        SET main_category = '총재고'
        WHERE main_category = '재고'
    """)

    conn.commit()
    conn.close()


def clean_field(field):
    return (
        field.replace("\x1e", "")
        .replace("\x04", "")
        .strip()
    )


def parse_barcode(raw):
    """
    실제 Data Matrix 규칙:
      - ALC는 GS(\x1d)로 분리된 S 필드에서만 읽음
      - SU... / SL... -> U... / L... -> FRT
      - SR... / SS... -> R... / S... -> RR
      - 예: SRA8 -> RA8 -> RR
      - V... 필드는 ALC로 사용하지 않음 (예: VR148 무시)
    """
    if not raw:
        return {
            "alc_code": "",
            "alc_type": "",
            "p_code": "",
            "part_number": "",
            "product_date": "",
        }

    records = raw.split("#")

    for record in records:
        fields = [clean_field(x) for x in record.split("\x1d")]

        record_alc = ""
        record_alc_type = ""
        record_p = ""
        record_part = ""
        record_date = ""

        for field in fields:
            if not field:
                continue

            # ALC: 반드시 S 필드만 사용
            if not record_alc and field.upper().startswith("S") and len(field) >= 2:
                candidate = field[1:].strip().upper()

                if re.fullmatch(r"[ULRS][A-Z0-9]+", candidate):
                    record_alc = candidate

                    if candidate[0] in ("U", "L"):
                        record_alc_type = "FRT"
                    elif candidate[0] in ("R", "S"):
                        record_alc_type = "RR"

            # P 코드
            if not record_p and field.startswith("P") and len(field) > 1:
                record_p = field[1:].strip()

            # 부품번호: CL4-... -> L4-...
            if not record_part and field.startswith("CL4-"):
                record_part = field[1:].strip()

            # 생산일자: T260730... -> 2026-07-30
            if not record_date and field.startswith("T") and len(field) >= 7:
                date_text = field[1:7]

                if date_text.isdigit():
                    year = "20" + date_text[0:2]
                    month = date_text[2:4]
                    day = date_text[4:6]

                    try:
                        parsed_date = datetime.strptime(
                            f"{year}-{month}-{day}",
                            "%Y-%m-%d",
                        )
                        record_date = parsed_date.strftime("%Y-%m-%d")
                    except ValueError:
                        pass

        # ALC가 있는 첫 레코드를 사용
        if record_alc:
            return {
                "alc_code": record_alc,
                "alc_type": record_alc_type,
                "p_code": record_p,
                "part_number": record_part,
                "product_date": record_date,
            }

    # ALC를 못 찾은 경우에도 부품번호/일자는 보조 검색
    part_number = ""
    product_date = ""

    part_match = re.search(r"L4-[A-Za-z0-9]+-\d+", raw)
    if part_match:
        part_number = part_match.group(0)

    date_match = re.search(
        r"(20\d{2})[-./](\d{1,2})[-./](\d{1,2})",
        raw,
    )
    if date_match:
        year = date_match.group(1)
        month = date_match.group(2).zfill(2)
        day = date_match.group(3).zfill(2)

        try:
            product_date = datetime.strptime(
                f"{year}-{month}-{day}",
                "%Y-%m-%d",
            ).strftime("%Y-%m-%d")
        except ValueError:
            pass

    return {
        "alc_code": "",
        "alc_type": "",
        "p_code": "",
        "part_number": part_number,
        "product_date": product_date,
    }


def count_query(conn, sql, params=()):
    return conn.execute(sql, params).fetchone()[0]


def get_statistics(conn):
    today = today_string()

    # 1. 당일생산분
    production_finished_frt = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '당일생산분'
          AND sub_category = '완제품'
          AND alc_type = 'FRT'
          AND DATE(scanned_at) = ?
        """,
        (today,),
    )

    production_finished_rr = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '당일생산분'
          AND sub_category = '완제품'
          AND alc_type = 'RR'
          AND DATE(scanned_at) = ?
        """,
        (today,),
    )

    production_semi_frt = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '당일생산분'
          AND sub_category = '반제품'
          AND alc_type = 'FRT'
          AND DATE(scanned_at) = ?
        """,
        (today,),
    )

    production_semi_rr = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '당일생산분'
          AND sub_category = '반제품'
          AND alc_type = 'RR'
          AND DATE(scanned_at) = ?
        """,
        (today,),
    )

    today_production = (
        production_finished_frt
        + production_finished_rr
        + production_semi_frt
        + production_semi_rr
    )

    # 2. 누적 총재고
    stock_finished_frt = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '총재고'
          AND sub_category = '완제품'
          AND alc_type = 'FRT'
        """,
    )

    stock_finished_rr = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '총재고'
          AND sub_category = '완제품'
          AND alc_type = 'RR'
        """,
    )

    stock_semi_frt = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '총재고'
          AND sub_category = '반제품'
          AND alc_type = 'FRT'
        """,
    )

    stock_semi_rr = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '총재고'
          AND sub_category = '반제품'
          AND alc_type = 'RR'
        """,
    )

    stock_frt = stock_finished_frt + stock_semi_frt
    stock_rr = stock_finished_rr + stock_semi_rr
    stock_total = stock_frt + stock_rr

    # 3. 누적 출고
    shipped_frt = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '출고'
          AND alc_type = 'FRT'
        """,
    )

    shipped_rr = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '출고'
          AND alc_type = 'RR'
        """,
    )

    shipped_total = shipped_frt + shipped_rr

    shipped_today_frt = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '출고'
          AND alc_type = 'FRT'
          AND DATE(scanned_at) = ?
        """,
        (today,),
    )

    shipped_today_rr = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '출고'
          AND alc_type = 'RR'
          AND DATE(scanned_at) = ?
        """,
        (today,),
    )

    # 4. 불량
    defective_frt = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '불량'
          AND alc_type = 'FRT'
        """,
    )

    defective_rr = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '불량'
          AND alc_type = 'RR'
        """,
    )

    defective_today_frt = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '불량'
          AND alc_type = 'FRT'
          AND DATE(scanned_at) = ?
        """,
        (today,),
    )

    defective_today_rr = count_query(
        conn,
        """
        SELECT COUNT(*) FROM scans
        WHERE main_category = '불량'
          AND alc_type = 'RR'
          AND DATE(scanned_at) = ?
        """,
        (today,),
    )

    # 합계 = 누적 총재고 - 누적 출고
    balance_frt = stock_frt - shipped_frt
    balance_rr = stock_rr - shipped_rr
    balance_total = stock_total - shipped_total

    alc_rows = conn.execute("""
        SELECT alc_type, alc_code, COUNT(*) AS count
        FROM scans
        WHERE alc_code IS NOT NULL
          AND alc_code != ''
        GROUP BY alc_type, alc_code
        ORDER BY alc_type, alc_code
    """).fetchall()

    today_alc_rows = conn.execute(
        """
        SELECT alc_type, alc_code, COUNT(*) AS count
        FROM scans
        WHERE alc_code IS NOT NULL
          AND alc_code != ''
          AND DATE(scanned_at) = ?
        GROUP BY alc_type, alc_code
        ORDER BY alc_type, alc_code
        """,
        (today,),
    ).fetchall()

    date_rows = conn.execute("""
        SELECT product_date, COUNT(*) AS count
        FROM scans
        WHERE product_date IS NOT NULL
          AND product_date != ''
        GROUP BY product_date
        ORDER BY product_date DESC
    """).fetchall()

    scan_rows = conn.execute("""
        SELECT *
        FROM scans
        ORDER BY id DESC
        LIMIT 30
    """).fetchall()

    return {
        "today_production": today_production,
        "production_finished_frt": production_finished_frt,
        "production_finished_rr": production_finished_rr,
        "production_semi_frt": production_semi_frt,
        "production_semi_rr": production_semi_rr,

        "stock_finished_frt": stock_finished_frt,
        "stock_finished_rr": stock_finished_rr,
        "stock_semi_frt": stock_semi_frt,
        "stock_semi_rr": stock_semi_rr,
        "stock_frt": stock_frt,
        "stock_rr": stock_rr,
        "stock_total": stock_total,

        "shipped_frt": shipped_frt,
        "shipped_rr": shipped_rr,
        "shipped_total": shipped_total,
        "shipped_today_frt": shipped_today_frt,
        "shipped_today_rr": shipped_today_rr,

        "defective_frt": defective_frt,
        "defective_rr": defective_rr,
        "defective_today_frt": defective_today_frt,
        "defective_today_rr": defective_today_rr,

        "balance_frt": balance_frt,
        "balance_rr": balance_rr,
        "balance_total": balance_total,

        "alc_counts": [dict(row) for row in alc_rows],
        "today_alc_counts": [dict(row) for row in today_alc_rows],
        "date_counts": [dict(row) for row in date_rows],
        "scans": [dict(row) for row in scan_rows],
    }


@app.after_request
def disable_cache(response):
    # 휴대폰 브라우저가 예전 index.html / JS를 계속 쓰는 문제 방지
    if request.path == "/":
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


@app.route("/")
def index():
    conn = get_db()
    stats = get_statistics(conn)
    conn.close()

    response = make_response(
        render_template(
            "index.html",
            **stats,
            page_version=now_kst().strftime("%Y%m%d%H%M%S"),
        )
    )
    return response


@app.route("/scan", methods=["POST"])
def scan():
    try:
        data = request.get_json(silent=True)

        if not data:
            return jsonify({
                "success": False,
                "error": "전송된 데이터가 없습니다.",
            }), 400

        raw = str(data.get("raw", ""))

        main_category = str(
            data.get("main_category", "")
        ).strip()

        sub_category = str(
            data.get("sub_category", "")
        ).strip()

        # 대소문자/공백 문제까지 방지
        selected_alc_type = str(
            data.get("selected_alc_type", "")
        ).strip().upper()

        print()
        print("==============================")
        print("브라우저 선택값")
        print("==============================")
        print("메인:", repr(main_category))
        print("서브:", repr(sub_category))
        print("선택 ALC:", repr(selected_alc_type))
        print("==============================")

        print("바코드 원본 데이터")
        print("==============================")
        print(repr(raw))
        print("==============================")

        if main_category not in (
            "당일생산분",
            "총재고",
            "출고",
            "불량",
        ):
            return jsonify({
                "success": False,
                "error": "올바른 카테고리를 선택해주세요.",
            }), 400

        if main_category in (
            "당일생산분",
            "총재고",
        ):
            if sub_category not in (
                "완제품",
                "반제품",
            ):
                return jsonify({
                    "success": False,
                    "error": "완제품 또는 반제품을 선택해주세요.",
                }), 400
        else:
            sub_category = ""

        if selected_alc_type not in ("FRT", "RR"):
            return jsonify({
                "success": False,
                "error": (
                    "FRT 또는 RR을 다시 선택해주세요.\n\n"
                    f"브라우저가 보낸 값: {selected_alc_type or '없음'}"
                ),
            }), 400

        if not raw:
            return jsonify({
                "success": False,
                "error": "바코드 데이터를 읽지 못했습니다.",
            }), 400

        parsed = parse_barcode(raw)

        alc_code = parsed["alc_code"]
        alc_type = parsed["alc_type"]
        p_code = parsed["p_code"]
        part_number = parsed["part_number"]
        product_date = parsed["product_date"]

        print("바코드 분석 결과")
        print("==============================")
        print("ALC:", alc_code)
        print("ALC 구분:", alc_type)
        print("P코드:", p_code)
        print("부품번호:", part_number)
        print("생산일자:", product_date)
        print("==============================")

        if not alc_code:
            return jsonify({
                "success": False,
                "error": "ALC 코드를 찾지 못했습니다.",
            }), 400

        # 서버가 인식한 실제 ALC와 사용자가 선택한 버튼을 최종 비교
        if selected_alc_type != alc_type:
            return jsonify({
                "success": False,
                "error": (
                    "ALC 구분이 일치하지 않습니다.\n\n"
                    f"선택한 구분: {selected_alc_type}\n"
                    f"실제 인식: {alc_type}\n"
                    f"인식된 ALC: {alc_code}\n\n"
                    "화면에서 올바른 FRT/RR 버튼을 다시 눌러주세요."
                ),
                "selected_alc_type": selected_alc_type,
                "detected_alc_type": alc_type,
                "detected_alc_code": alc_code,
            }), 400

        scanned_at = now_kst().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conn = get_db()

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
            scanned_at,
        ))

        conn.commit()

        alc_total = count_query(
            conn,
            """
            SELECT COUNT(*)
            FROM scans
            WHERE alc_code = ?
            """,
            (alc_code,),
        )

        alc_today = count_query(
            conn,
            """
            SELECT COUNT(*)
            FROM scans
            WHERE alc_code = ?
              AND DATE(scanned_at) = ?
            """,
            (
                alc_code,
                today_string(),
            ),
        )

        stats = get_statistics(conn)
        conn.close()

        response_data = {
            "success": True,
            "main_category": main_category,
            "sub_category": sub_category,
            "selected_alc_type": selected_alc_type,
            "alc_code": alc_code,
            "alc_type": alc_type,
            "alc_total": alc_total,
            "alc_today": alc_today,
            "p_code": p_code,
            "part_number": part_number,
            "product_date": product_date,
            "scanned_at": scanned_at,
        }

        response_data.update(stats)

        print("저장 완료")
        print("==============================")
        print("업무:", main_category)
        print("제품:", sub_category)
        print("선택:", selected_alc_type)
        print("실제:", alc_type)
        print("ALC:", alc_code)
        print("==============================")

        return jsonify(response_data)

    except Exception as e:
        print("오류 발생:", str(e))

        return jsonify({
            "success": False,
            "error": str(e),
        }), 500


@app.route("/export.xlsx")
def export_excel():
    conn = get_db()
    stats = get_statistics(conn)

    all_scans = conn.execute("""
        SELECT *
        FROM scans
        ORDER BY id ASC
    """).fetchall()

    daily_rows = conn.execute("""
        SELECT
            DATE(scanned_at) AS date,

            SUM(
                CASE
                    WHEN main_category = '총재고'
                    THEN 1
                    ELSE 0
                END
            ) AS stock_added,

            SUM(
                CASE
                    WHEN main_category = '출고'
                    THEN 1
                    ELSE 0
                END
            ) AS shipped,

            SUM(
                CASE
                    WHEN main_category = '불량'
                    THEN 1
                    ELSE 0
                END
            ) AS defective

        FROM scans
        WHERE scanned_at IS NOT NULL
          AND scanned_at != ''

        GROUP BY DATE(scanned_at)
        ORDER BY DATE(scanned_at)
    """).fetchall()

    conn.close()

    wb = Workbook()

    header_fill = PatternFill(
        "solid",
        fgColor="D9EAF7",
    )
    header_font = Font(
        bold=True
    )

    # 1. 요약
    ws = wb.active
    ws.title = "요약"

    ws.append([
        "항목",
        "FRT",
        "RR",
        "합계",
    ])

    ws.append([
        "누적 총재고",
        stats["stock_frt"],
        stats["stock_rr"],
        stats["stock_total"],
    ])

    ws.append([
        "누적 출고",
        stats["shipped_frt"],
        stats["shipped_rr"],
        stats["shipped_total"],
    ])

    ws.append([
        "현재 재고",
        stats["balance_frt"],
        stats["balance_rr"],
        stats["balance_total"],
    ])

    ws.append([
        "누적 불량",
        stats["defective_frt"],
        stats["defective_rr"],
        stats["defective_frt"] + stats["defective_rr"],
    ])

    ws.append([])

    ws.append([
        "총재고 세부",
        "FRT",
        "RR",
        "합계",
    ])

    ws.append([
        "완제품",
        stats["stock_finished_frt"],
        stats["stock_finished_rr"],
        (
            stats["stock_finished_frt"]
            + stats["stock_finished_rr"]
        ),
    ])

    ws.append([
        "반제품",
        stats["stock_semi_frt"],
        stats["stock_semi_rr"],
        (
            stats["stock_semi_frt"]
            + stats["stock_semi_rr"]
        ),
    ])

    # 2. 일자별 재고
    ws_daily = wb.create_sheet(
        "일자별 재고"
    )

    ws_daily.append([
        "날짜",
        "당일 재고 추가",
        "당일 출고",
        "당일 불량",
        "누적 총재고",
        "누적 출고",
        "현재 재고",
    ])

    cumulative_stock = 0
    cumulative_shipped = 0

    for row in daily_rows:
        stock_added = row["stock_added"] or 0
        shipped = row["shipped"] or 0
        defective = row["defective"] or 0

        cumulative_stock += stock_added
        cumulative_shipped += shipped

        current_balance = (
            cumulative_stock
            - cumulative_shipped
        )

        ws_daily.append([
            row["date"],
            stock_added,
            shipped,
            defective,
            cumulative_stock,
            cumulative_shipped,
            current_balance,
        ])

    # 3. ALC 수량
    ws_alc = wb.create_sheet(
        "ALC 수량"
    )

    ws_alc.append([
        "ALC 구분",
        "ALC 코드",
        "수량",
    ])

    for row in stats["alc_counts"]:
        ws_alc.append([
            row["alc_type"],
            row["alc_code"],
            row["count"],
        ])

    # 4. 전체 스캔 기록
    ws_scan = wb.create_sheet(
        "전체 스캔 기록"
    )

    ws_scan.append([
        "번호",
        "업무구분",
        "제품종류",
        "ALC구분",
        "ALC",
        "P코드",
        "부품번호",
        "생산일자",
        "스캔시간",
    ])

    for row in all_scans:
        ws_scan.append([
            row["id"],
            row["main_category"],
            row["sub_category"],
            row["alc_type"],
            row["alc_code"],
            row["p_code"],
            row["part_number"],
            row["product_date"],
            row["scanned_at"],
        ])

    for sheet in wb.worksheets:
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(
                horizontal="center"
            )

        for column in sheet.columns:
            max_length = 0
            column_letter = (
                column[0].column_letter
            )

            for cell in column:
                value = (
                    ""
                    if cell.value is None
                    else str(cell.value)
                )

                max_length = max(
                    max_length,
                    len(value),
                )

            sheet.column_dimensions[
                column_letter
            ].width = min(
                max_length + 3,
                35,
            )

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = (
        "barcode_report_"
        + now_kst().strftime(
            "%Y%m%d_%H%M%S"
        )
        + ".xlsx"
    )

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


init_db()


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True,
    )

import streamlit as st
import pandas as pd
import re
from pathlib import Path

# ==========================================
# WAREHOUSE AI - TRA VỊ TRÍ HÀNG
# ==========================================

st.set_page_config(
    page_title="WAREHOUSE AI",
    page_icon="📦",
    layout="centered"
)

st.title("📦 WAREHOUSE AI")
st.caption("Hỏi vị trí hàng trong kho")

SHEET_NAME = "1804"


# ==========================================
# TỰ ĐỘNG TÌM FILE EXCEL MỚI NHẤT
# ==========================================

def get_latest_excel():

    folder = Path(__file__).parent

    files = [
        f for f in folder.glob("*.xlsx")
        if not f.name.startswith("~$")
    ]

    if not files:
        return None

    return max(
        files,
        key=lambda f: f.stat().st_mtime
    )


excel_file = get_latest_excel()

if excel_file is None:
    st.error(
        "Không tìm thấy file Excel trong thư mục WAREHOUSE_AI."
    )
    st.stop()


# ==========================================
# ĐỌC SHEET 1804
# ==========================================

@st.cache_data
def load_data(file_path, modified_time):

    df = pd.read_excel(
        file_path,
        sheet_name=SHEET_NAME
    )

    for col in df.columns:

        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
        )

    return df


try:

    df = load_data(
        str(excel_file),
        excel_file.stat().st_mtime
    )

except Exception as e:

    st.error(
        f"Không đọc được file Excel: {e}"
    )

    st.stop()


# Hiển thị file dữ liệu đang sử dụng
st.caption(
    f"📊 Dữ liệu: {excel_file.name}"
)


# ==========================================
# TÌM CỘT SAP
# ==========================================

def find_column(keywords):

    for col in df.columns:

        name = str(col).upper().strip()

        if all(
            keyword.upper() in name
            for keyword in keywords
        ):
            return col

    return None


bin_col = find_column(
    ["STORAGE", "BIN"]
)

desc_col = find_column(
    ["MATERIAL", "DESCRIPTION"]
)


if bin_col is None:

    st.error(
        "Không tìm thấy cột Storage Bin."
    )
    st.stop()


if desc_col is None:

    st.error(
        "Không tìm thấy cột Material Description."
    )
    st.stop()


# ==========================================
# CHUẨN HÓA CÂU HỎI
# ==========================================

def normalize(text):

    text = str(text).upper().strip()

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text


# ==========================================
# ĐỌC CÂU HỎI
# ==========================================

def parse_question(question):

    q = normalize(question)

    # ------------------------------
    # CHỦNG LOẠI
    # ------------------------------

    product = None

    if re.search(r"\bMDF\b", q):
        product = "MDF"

    elif re.search(r"\bMFC\b", q):
        product = "MFC"

    elif re.search(r"\bHDF\b", q):
        product = "HDF"


    # ------------------------------
    # MUF / UF
    # ------------------------------

    glue = None

    if re.search(r"\bMUF\b", q):
        glue = "MUF"

    elif re.search(r"\bUF\b", q):
        glue = "UF"


    # ------------------------------
    # ĐỘ DÀY
    # ------------------------------

    thickness = None

    match = re.search(
        r"\b(\d{1,2})\s*(?:LY|MM)\b",
        q
    )

    if match:

        thickness = int(
            match.group(1)
        )

    else:

        # Cho phép hỏi:
        # MDF 388EV 17 MUF

        numbers = re.findall(
            r"\b(\d{1,2})\b",
            q
        )

        valid_thickness = {
            6,
            9,
            10,
            12,
            15,
            16,
            17,
            18,
            25
        }

        for number in numbers:

            number = int(number)

            if number in valid_thickness:

                thickness = number
                break


    # ------------------------------
    # MÃ MÀU
    # ------------------------------

    color = None

    ignored_words = {
        "MDF",
        "MFC",
        "HDF",
        "MUF",
        "UF",
        "BIN",
        "LY",
        "MM"
    }

    tokens = re.findall(
        r"\b[A-Z0-9-]+\b",
        q
    )

    for token in tokens:

        if token in ignored_words:
            continue

        if (
            re.search(r"\d", token)
            and
            re.search(r"[A-Z]", token)
        ):

            # Không lấy nhầm mã BIN
            if not re.match(
                r"^\d{3}-",
                token
            ):

                color = token
                break


    return (
        product,
        color,
        thickness,
        glue
    )


# ==========================================
# TÌM BIN
# ==========================================

def search_bins(question):

    (
        product,
        color,
        thickness,
        glue
    ) = parse_question(question)


    # ------------------------------
    # KIỂM TRA THÔNG TIN
    # ------------------------------

    missing = []

    if not product:
        missing.append(
            "chủng loại MDF/MFC/HDF"
        )

    if not color:
        missing.append(
            "mã màu"
        )

    if not thickness:
        missing.append(
            "độ dày"
        )

    if not glue:
        missing.append(
            "MUF/UF"
        )


    if missing:

        return None, (
            "Vui lòng bổ sung: "
            + ", ".join(missing)
        )


    result = df.copy()


    # ======================================
    # LỌC CHỦNG LOẠI
    # ======================================

    description = (
        result[desc_col]
        .str.upper()
    )


    if product == "MDF":

        mask_product = (
            description.str.contains(
                "MELMDF",
                regex=False,
                na=False
            )
            |
            description.str.contains(
                r"\bMDF\b",
                regex=True,
                na=False
            )
        )


    elif product == "MFC":

        mask_product = (
            description.str.contains(
                r"\bMFC\b",
                regex=True,
                na=False
            )
        )


    elif product == "HDF":

        mask_product = (
            description.str.contains(
                r"\bHDF\b",
                regex=True,
                na=False
            )
        )


    else:

        mask_product = pd.Series(
            False,
            index=result.index
        )


    result = result[
        mask_product
    ]


    # ======================================
    # LỌC MÃ MÀU
    # ======================================

    description = (
        result[desc_col]
        .str.upper()
    )

    result = result[
        description.str.contains(
            re.escape(color),
            regex=True,
            na=False
        )
    ]


    # ======================================
    # LỌC ĐỘ DÀY
    #
    # 9ly  -> 4809
    # 17ly -> 4817
    # 18ly -> 4818
    # ======================================

    thickness_code = (
        f"48{thickness:02d}"
    )

    description = (
        result[desc_col]
        .str.upper()
    )

    result = result[
        description.str.contains(
            thickness_code,
            regex=False,
            na=False
        )
    ]


    # ======================================
    # LỌC MUF / UF
    #
    # MUFSTD, MUFE1...
    # UFSTD, UFE1...
    #
    # UF KHÔNG ĐƯỢC ĂN NHẦM MUF
    # ======================================

    description = (
        result[desc_col]
        .str.upper()
    )


    if glue == "MUF":

        result = result[
            description.str.contains(
                "MUF",
                regex=False,
                na=False
            )
        ]


    elif glue == "UF":

        result = result[
            description.str.contains(
                "UF",
                regex=False,
                na=False
            )
            &
            ~description.str.contains(
                "MUF",
                regex=False,
                na=False
            )
        ]


    # ======================================
    # LẤY BIN ĐẦY ĐỦ
    #
    # 184-A02 và 185-A02
    # là 2 BIN KHÁC NHAU
    # ======================================

    bins = (
        result[bin_col]
        .astype(str)
        .str.strip()
    )


    bins = [

        bin_name

        for bin_name in bins.unique()

        if (
            bin_name
            and
            bin_name.upper() != "NAN"
        )
    ]


    bins = sorted(
        bins
    )


    info = {

        "product": product,

        "color": color,

        "thickness": thickness,

        "glue": glue,

        "bins": bins,

        "count": len(result)

    }


    return info, None


# ==========================================
# GIAO DIỆN CHAT
# ==========================================

question = st.chat_input(
    "Ví dụ: MDF 388EV 17ly MUF nằm ở BIN nào?"
)


if question:

    with st.chat_message("user"):

        st.write(
            question
        )


    info, error = search_bins(
        question
    )


    with st.chat_message("assistant"):


        if error:

            st.warning(
                error
            )


        elif not info["bins"]:

            st.error(

                f"Không tìm thấy vị trí cho "

                f"{info['product']} "

                f"{info['color']} "

                f"{info['thickness']}ly "

                f"{info['glue']}."

            )


        else:

            st.success(

                f"{info['product']} "

                f"{info['color']} "

                f"{info['thickness']}ly "

                f"{info['glue']} đang nằm tại:"

            )

        for bin_name in info["bins"]:
            st.markdown(
                f"### 📍 {bin_name}"
            )

        st.caption(
            f"Tìm thấy {len(info['bins'])} BIN."
        )

from datetime import datetime

_CH = ["không","một","hai","ba","bốn","năm","sáu","bảy","tám","chín"]
_DVN = ["","nghìn","triệu","tỷ"]

def _read_triple(n: int, is_first: bool = True) -> str:
    """Đọc một nhóm 3 chữ số (0-999)."""
    if n == 0:
        return ""
    tram = n // 100
    chuc = (n % 100) // 10
    don  = n % 10
    parts = []
    if tram > 0:
        parts.append(_CH[tram] + " trăm")
        if chuc == 0 and don > 0:
            parts.append("lẻ")
    elif not is_first:
        if chuc == 0 and don > 0:
            parts.append("không trăm lẻ")
        elif chuc > 0:
            parts.append("không trăm")
    if chuc == 1:
        parts.append("mười")
        if don > 0:
            parts.append(_CH[don])
    elif chuc > 1:
        parts.append(_CH[chuc] + " mươi")
        if don == 1:
            parts.append("mốt")
        elif don == 5:
            parts.append("lăm")
        elif don > 0:
            parts.append(_CH[don])
    elif don > 0 and (tram > 0 or not is_first):
        parts.append(_CH[don])
    elif don > 0:
        parts.append(_CH[don])
    return " ".join(parts)

def so_tien_bang_chu(amount: int) -> str:
    """Chuyển số tiền (nguyên, VND) thành chữ tiếng Việt, viết hoa chữ đầu."""
    if amount == 0:
        return "Không đồng chẵn"
    groups = []
    n = amount
    while n > 0:
        groups.append(n % 1000)
        n //= 1000
    parts = []
    for i in range(len(groups) - 1, -1, -1):
        g = groups[i]
        if g == 0:
            continue
        text = _read_triple(g, is_first=(i == len(groups) - 1))
        if _DVN[i]:
            text += " " + _DVN[i]
        parts.append(text)
    result = " ".join(parts).strip() + " đồng chẵn"
    # Viết hoa chữ đầu
    return result[0].upper() + result[1:] if result else result

def format_date(date_str: str) -> str:
    """'2026-03-27' → 'ngày 27 tháng 03 năm 2026'"""
    try:
        d = datetime.strptime(date_str, "%Y-%m-%d")
        return f"ngày {d.day:02d} tháng {d.month:02d} năm {d.year}"
    except Exception:
        return date_str

def format_currency(amount: int) -> str:
    """1166080 → '1.166.080'"""
    return f"{amount:,}".replace(",", ".")

"""
蝦皮訂單 & 廣告費分析工具
讓蝦皮賣家可以用自己的數據，分析營業額、廣告費、優惠券、平台抽成的趨勢。

使用方式：
1. 從蝦皮後台匯出「訂單」Excel 檔案（可加密）
2. 從蝦皮後台匯出「廣告數據」CSV 檔案
3. 將檔案分別放在 "訂單" 和 "廣告費" 資料夾
4. 執行本程式，依提示輸入路徑與密碼
"""

import os
import sys
import io
import glob
import re
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("錯誤：缺少 pandas 套件，請執行 pip install pandas")
    input("按 Enter 結束...")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("錯誤：缺少 openpyxl 套件，請執行 pip install openpyxl")
    input("按 Enter 結束...")
    sys.exit(1)

try:
    import msoffcrypto
except ImportError:
    msoffcrypto = None


def open_encrypted_xlsx(path, password):
    """解密加密的 xlsx 檔案"""
    if msoffcrypto is None:
        print("錯誤：檔案有密碼保護，但缺少 msoffcrypto 套件。")
        print("請執行：pip install msoffcrypto-tool")
        input("按 Enter 結束...")
        sys.exit(1)
    with open(path, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=password)
        decrypted = io.BytesIO()
        file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted


def load_order_file(path, password=None):
    """載入單一訂單 Excel 檔案"""
    try:
        wb = openpyxl.load_workbook(path, read_only=False)
    except Exception:
        if password:
            dec = open_encrypted_xlsx(path, password)
            wb = openpyxl.load_workbook(dec, read_only=False)
        else:
            print(f"  無法開啟 {path}，檔案可能有密碼保護。")
            return None

    ws = wb.active
    headers = [ws.cell(row=1, column=i).value for i in range(1, ws.max_column + 1)]
    data = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        data.append(row)
    wb.close()
    return pd.DataFrame(data, columns=headers)


def load_ad_file(path):
    """載入單一廣告 CSV 檔案"""
    for enc in ['utf-8-sig', 'utf-8', 'big5', 'cp950']:
        try:
            df = pd.read_csv(path, skiprows=7, encoding=enc)
            if '花費' in df.columns:
                return df
        except Exception:
            continue
    print(f"  無法讀取廣告檔案：{path}")
    return None


def to_num(series):
    """將欄位轉為數值"""
    return pd.to_numeric(series.astype(str).str.replace(',', '').str.replace('%', ''), errors='coerce')


def extract_month_label(filename):
    """從檔名提取月份標籤"""
    # 訂單檔: Order.all.20251201_20251231.xlsx
    m = re.search(r'(\d{4})(\d{2})\d{2}_\d{8}', filename)
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        return f"{year}/{month}月", year * 100 + month

    # 廣告檔: 蝦皮廣告-總體-數據-2025_12_01-2025_12_31.csv
    m = re.search(r'(\d{4})_(\d{2})_\d{2}-\d{4}_\d{2}_\d{2}', filename)
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        return f"{year}/{month}月", year * 100 + month

    return os.path.basename(filename), 0


def analyze_orders(order_files, password=None):
    """分析所有訂單檔案"""
    results = {}

    for path in order_files:
        label, sort_key = extract_month_label(os.path.basename(path))
        print(f"  讀取訂單：{os.path.basename(path)} → {label}")

        df = load_order_file(path, password)
        if df is None:
            continue

        # 必要欄位檢查
        required = ['訂單編號', '商品總價', '買家總支付金額', '優惠代碼',
                     '蝦皮負擔優惠券', '蝦幣折抵', '成交手續費', '其他服務費', '金流與系統處理費']
        missing = [c for c in required if c not in df.columns]
        if missing:
            print(f"  警告：缺少欄位 {missing}，跳過此檔案")
            continue

        u = df.drop_duplicates(subset='訂單編號', keep='first')

        revenue = pd.to_numeric(u['商品總價'], errors='coerce').sum()
        buyer_paid = pd.to_numeric(u['買家總支付金額'], errors='coerce').sum()
        orders = len(u)

        has_coupon = u[u['優惠代碼'].notna() & (u['優惠代碼'] != '')]
        coupon_orders = len(has_coupon)

        shopee_coupon = pd.to_numeric(u['蝦皮負擔優惠券'], errors='coerce').sum()
        coin_discount = pd.to_numeric(u['蝦幣折抵'], errors='coerce').sum()
        total_discount = revenue - buyer_paid

        tx_fee = pd.to_numeric(u['成交手續費'], errors='coerce').sum()
        other_fee = pd.to_numeric(u['其他服務費'], errors='coerce').sum()
        payment_fee = pd.to_numeric(u['金流與系統處理費'], errors='coerce').sum()
        platform_total = tx_fee + other_fee + payment_fee

        results[label] = {
            'sort_key': sort_key,
            'orders': orders,
            'coupon_orders': coupon_orders,
            'revenue': revenue,
            'buyer_paid': buyer_paid,
            'total_discount': total_discount,
            'shopee_coupon': shopee_coupon,
            'coin_discount': coin_discount,
            'tx_fee': tx_fee,
            'other_fee': other_fee,
            'payment_fee': payment_fee,
            'platform_total': platform_total,
        }

    return dict(sorted(results.items(), key=lambda x: x[1]['sort_key']))


def analyze_ads(ad_files):
    """分析所有廣告檔案"""
    results = {}

    for path in ad_files:
        label, sort_key = extract_month_label(os.path.basename(path))
        print(f"  讀取廣告：{os.path.basename(path)} → {label}")

        df = load_ad_file(path)
        if df is None:
            continue

        ad_spend = to_num(df['花費']).sum()
        ad_sales = to_num(df['銷售金額']).sum()
        clicks = to_num(df['點擊數']).sum()
        impressions = to_num(df['瀏覽數']).sum()
        conversions = to_num(df['轉換數']).sum()

        roas = ad_sales / ad_spend if ad_spend > 0 else 0
        cpc = ad_spend / clicks if clicks > 0 else 0
        conv_rate = conversions / clicks * 100 if clicks > 0 else 0

        results[label] = {
            'sort_key': sort_key,
            'ad_spend': ad_spend,
            'ad_sales': ad_sales,
            'clicks': clicks,
            'impressions': impressions,
            'conversions': conversions,
            'roas': roas,
            'cpc': cpc,
            'conv_rate': conv_rate,
        }

    return dict(sorted(results.items(), key=lambda x: x[1]['sort_key']))


def print_report(order_results, ad_results):
    """輸出分析報告"""
    # 合併月份
    all_months = list(dict.fromkeys(list(order_results.keys()) + list(ad_results.keys())))

    col_width = 12
    header = f"{'':30s}" + "".join(f"{m:>{col_width}s}" for m in all_months)
    sep = "─" * (30 + col_width * len(all_months))

    print("\n" + "=" * len(sep))
    print("【蝦皮賣場分析報告】")
    print("=" * len(sep))

    # === 訂單數據 ===
    if order_results:
        print(f"\n{header}")
        print(sep)

        def pval(key, fmt=',.0f', prefix='$'):
            vals = []
            for m in all_months:
                if m in order_results:
                    v = order_results[m][key]
                    if prefix == '$':
                        vals.append(f"${v:>{col_width-1}{fmt}}")
                    else:
                        vals.append(f"{v:>{col_width}{fmt}}")
                else:
                    vals.append(f"{'N/A':>{col_width}s}")
            return "".join(vals)

        print(f"  {'訂單數':28s}{pval('orders', 'd', '')}")
        print(f"  {'有優惠券訂單數':28s}{pval('coupon_orders', 'd', '')}")

        # 優惠券使用率
        rate_vals = []
        for m in all_months:
            if m in order_results:
                r = order_results[m]
                rate = r['coupon_orders'] / r['orders'] * 100 if r['orders'] > 0 else 0
                rate_vals.append(f"{rate:>{col_width-1}.1f}%")
            else:
                rate_vals.append(f"{'N/A':>{col_width}s}")
        print(f"  {'優惠券使用率':28s}{''.join(rate_vals)}")

        print(sep)
        print(f"  {'商品總價（營業額）':28s}{pval('revenue')}")
        print(f"  {'買家總支付金額':28s}{pval('buyer_paid')}")
        print(f"  {'所有折扣合計':28s}{pval('total_discount')}")
        print(f"  {'  └ 蝦皮負擔優惠券':28s}{pval('shopee_coupon')}")
        print(f"  {'  └ 蝦幣折抵':28s}{pval('coin_discount')}")

        print(sep)
        print(f"  {'成交手續費':28s}{pval('tx_fee')}")
        print(f"  {'其他服務費':28s}{pval('other_fee')}")
        print(f"  {'金流處理費':28s}{pval('payment_fee')}")
        print(f"  {'平台費用小計':28s}{pval('platform_total')}")

    # === 廣告數據 ===
    if ad_results:
        print(sep)
        ad_vals = []
        for m in all_months:
            if m in ad_results:
                v = ad_results[m]['ad_spend']
                ad_vals.append(f"${v:>{col_width-1},.0f}")
            else:
                ad_vals.append(f"{'N/A':>{col_width}s}")
        print(f"  {'廣告費':28s}{''.join(ad_vals)}")

    # === 總成本與到手 ===
    if order_results and ad_results:
        print(sep)
        cost_vals = []
        net_vals = []
        for m in all_months:
            if m in order_results and m in ad_results:
                o = order_results[m]
                a = ad_results[m]
                total_cost = o['platform_total'] + a['ad_spend']
                net = o['buyer_paid'] - o['platform_total'] - a['ad_spend']
                cost_vals.append(f"${total_cost:>{col_width-1},.0f}")
                net_vals.append(f"${net:>{col_width-1},.0f}")
            else:
                cost_vals.append(f"{'N/A':>{col_width}s}")
                net_vals.append(f"{'N/A':>{col_width}s}")
        print(f"  {'總成本(平台+廣告)':28s}{''.join(cost_vals)}")
        print(f"  {'實際到手估算':28s}{''.join(net_vals)}")

    # === 費率 ===
    print(f"\n{sep}")
    print(f"  {'【費率分析】':28s}")
    print(sep)

    if order_results:
        rate_line = []
        for m in all_months:
            if m in order_results:
                o = order_results[m]
                r = o['platform_total'] / o['revenue'] * 100 if o['revenue'] > 0 else 0
                rate_line.append(f"{r:>{col_width-1}.1f}%")
            else:
                rate_line.append(f"{'N/A':>{col_width}s}")
        print(f"  {'平台費率':28s}{''.join(rate_line)}")

    if ad_results and order_results:
        rate_line = []
        for m in all_months:
            if m in order_results and m in ad_results:
                r = ad_results[m]['ad_spend'] / order_results[m]['revenue'] * 100 if order_results[m]['revenue'] > 0 else 0
                rate_line.append(f"{r:>{col_width-1}.1f}%")
            else:
                rate_line.append(f"{'N/A':>{col_width}s}")
        print(f"  {'廣告費率(廣告/營收)':28s}{''.join(rate_line)}")

        rate_line = []
        for m in all_months:
            if m in order_results and m in ad_results:
                o = order_results[m]
                total = (o['platform_total'] + ad_results[m]['ad_spend']) / o['revenue'] * 100 if o['revenue'] > 0 else 0
                rate_line.append(f"{total:>{col_width-1}.1f}%")
            else:
                rate_line.append(f"{'N/A':>{col_width}s}")
        print(f"  {'總成本率':28s}{''.join(rate_line)}")

    # === 廣告效率 ===
    if ad_results:
        print(f"\n{sep}")
        print(f"  {'【廣告效率】':28s}")
        print(sep)

        for label, key, fmt, prefix in [
            ('ROAS', 'roas', '.1f', ''),
            ('CPC(每次點擊成本)', 'cpc', '.2f', '$'),
            ('廣告轉換率', 'conv_rate', '.1f', '%'),
        ]:
            vals = []
            for m in all_months:
                if m in ad_results:
                    v = ad_results[m][key]
                    if prefix == '$':
                        vals.append(f"${v:>{col_width-1}{fmt}}")
                    elif prefix == '%':
                        vals.append(f"{v:>{col_width-1}{fmt}}%")
                    else:
                        vals.append(f"{v:>{col_width}{fmt}}")
                else:
                    vals.append(f"{'N/A':>{col_width}s}")
            print(f"  {label:28s}{''.join(vals)}")

    # === 月度變化 ===
    months_with_both = [m for m in all_months if m in order_results and m in ad_results]
    if len(months_with_both) >= 2:
        base_month = months_with_both[0]
        bo = order_results[base_month]
        ba = ad_results[base_month]

        print(f"\n{sep}")
        print(f"  {'【月度變化】以 ' + base_month + ' 為基準':28s}")
        print(sep)

        for label, get_val in [
            ('營業額變化', lambda m: order_results[m]['revenue'] / bo['revenue'] * 100 - 100),
            ('訂單數變化', lambda m: order_results[m]['orders'] / bo['orders'] * 100 - 100),
            ('廣告費變化', lambda m: ad_results[m]['ad_spend'] / ba['ad_spend'] * 100 - 100),
        ]:
            vals = []
            for m in months_with_both:
                if m == base_month:
                    vals.append(f"{'基準':>{col_width}s}")
                else:
                    v = get_val(m)
                    vals.append(f"{v:>{col_width-1}+.1f}%")
            # Only print columns that have data
            print(f"  {label:28s}{''.join(vals)}")

        # 到手變化
        vals = []
        for m in months_with_both:
            base_net = bo['buyer_paid'] - bo['platform_total'] - ba['ad_spend']
            cur_o = order_results[m]
            cur_net = cur_o['buyer_paid'] - cur_o['platform_total'] - ad_results[m]['ad_spend']
            if m == base_month:
                vals.append(f"{'基準':>{col_width}s}")
            else:
                v = cur_net / base_net * 100 - 100
                vals.append(f"{v:>{col_width-1}+.1f}%")
        print(f"  {'實際到手變化':28s}{''.join(vals)}")

    print("\n" + "=" * len(sep))
    print("報告生成完畢")
    print("=" * len(sep))


def find_files(base_dir):
    """自動尋找訂單和廣告檔案"""
    order_dir = os.path.join(base_dir, "訂單")
    ad_dir = os.path.join(base_dir, "廣告費")

    order_files = []
    ad_files = []

    if os.path.isdir(order_dir):
        order_files = sorted(glob.glob(os.path.join(order_dir, "*.xlsx")))
    if os.path.isdir(ad_dir):
        ad_files = sorted(glob.glob(os.path.join(ad_dir, "*.csv")))

    return order_files, ad_files


def main():
    print("=" * 60)
    print("  蝦皮訂單 & 廣告費分析工具")
    print("  讓數據說話，拒絕黑箱")
    print("=" * 60)

    # 取得資料夾路徑
    print("\n請輸入資料夾路徑（內含「訂單」和「廣告費」子資料夾）")
    print("（直接按 Enter 使用目前目錄）")
    base_dir = input("路徑：").strip().strip('"').strip("'")
    if not base_dir:
        base_dir = os.getcwd()

    if not os.path.isdir(base_dir):
        print(f"錯誤：找不到資料夾 {base_dir}")
        input("按 Enter 結束...")
        return

    order_files, ad_files = find_files(base_dir)

    print(f"\n找到 {len(order_files)} 個訂單檔案, {len(ad_files)} 個廣告檔案")

    if not order_files and not ad_files:
        print("錯誤：找不到任何檔案。")
        print(f"請確認 {base_dir} 下有「訂單」和「廣告費」子資料夾。")
        input("按 Enter 結束...")
        return

    # 密碼
    password = None
    if order_files:
        print("\n訂單檔案是否有密碼保護？")
        pwd_input = input("密碼（無密碼請直接按 Enter）：").strip()
        if pwd_input:
            password = pwd_input

    # 分析
    print("\n正在分析訂單資料...")
    order_results = analyze_orders(order_files, password) if order_files else {}

    print("\n正在分析廣告資料...")
    ad_results = analyze_ads(ad_files) if ad_files else {}

    # 輸出報告
    print_report(order_results, ad_results)

    # 儲存報告到檔案
    report_path = os.path.join(base_dir, f"分析報告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    try:
        original_stdout = sys.stdout
        with open(report_path, 'w', encoding='utf-8') as f:
            sys.stdout = f
            print_report(order_results, ad_results)
        sys.stdout = original_stdout
        print(f"\n報告已儲存至：{report_path}")
    except Exception as e:
        print(f"\n儲存報告時發生錯誤：{e}")

    input("\n按 Enter 結束...")


if __name__ == "__main__":
    main()

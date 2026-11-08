import pandas as pd
import os
import shutil
import uuid
import traceback
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, Form, Body
from fastapi.responses import FileResponse, HTMLResponse
from pathlib import Path
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
from io import BytesIO

app = FastAPI(title="智能记账助手")

CANCEL_OUT_DAYS = 7

DEFAULT_FALLBACK_RULES = """
食=餐饮美食
餐=餐饮美食
饭=餐饮美食
面=餐饮美食
粉=餐饮美食
粥=餐饮美食
蛋糕=餐饮美食
面包=餐饮美食
奶茶=餐饮美食
咖啡=餐饮美食
汉堡=餐饮美食
炸鸡=餐饮美食
烧烤=餐饮美食
麻辣烫=餐饮美食
小吃=餐饮美食
零食=餐饮美食
水果=餐饮美食
外卖=餐饮美食
美团=餐饮美食
饿了么=餐饮美食
衣=服饰鞋包
裤=服饰鞋包
鞋=服饰鞋包
袜=服饰鞋包
帽=服饰鞋包
滴滴=交通出行
公交=交通出行
地铁=交通出行
打车=交通出行
高铁=交通出行
机票=交通出行
超市=日用百货
便利店=日用百货
药店=医疗健康
话费=通讯缴费
电费=生活缴费
水费=生活缴费
淘宝=网购
京东=网购
拼多多=网购
"""

VAGUE_CATEGORIES = ["其他", "转账", "未知", "商户消费", "扫二维码付款", "群收款", "微信红包", "nan", ""]


def clean_amount(val):
    if pd.isna(val): return 0.0
    s = str(val).replace("¥", "").replace("`", "").replace(",", "").replace('"', '').strip()
    if s.startswith('-'): s = s[1:]
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_user_rules(rule_text):
    combo_rules = []
    expand_dict = {}
    if not rule_text: return combo_rules, expand_dict
    for line in rule_text.strip().split('\n'):
        line = line.strip()
        if not line or line.startswith("#"): continue
        if "=" in line:
            parts = line.split("=")
            keyword_part = parts[0].strip()
            category = parts[1].strip()
            if "+" in keyword_part:
                targets = [t.strip() for t in keyword_part.split("+")]
                combo_rules.append({"targets": targets, "category": category})
            else:
                expand_dict[keyword_part] = category
    return combo_rules, expand_dict


def parse_file_to_df(filepath, filename):
    ext = Path(filepath).suffix.lower()

    if ext in ['.xlsx', '.xls']:
        try:
            df_raw = pd.read_excel(filepath, header=None)
            header_row = -1
            for i in range(min(50, len(df_raw))):
                row_str = str(df_raw.iloc[i].values)
                if "交易时间" in row_str or "交易日期" in row_str:
                    header_row = i
                    break
            if header_row != -1:
                df = pd.read_excel(filepath, header=header_row)
                df.columns = [str(c).strip() for c in df.columns]
                if len(df.columns) > 3:
                    return df
        except Exception:
            pass

    for encoding in ["utf-8-sig", "gbk", "gb18030", "utf-8"]:
        try:
            with open(filepath, 'r', encoding=encoding) as f:
                content = f.read()

            lines = content.splitlines()
            header_index = -1
            for i, line in enumerate(lines):
                if "交易时间" in line or "交易日期" in line:
                    header_index = i
                    break

            if header_index == -1: continue

            header_line = lines[header_index]
            delimiter = '\t' if '\t' in header_line else ','

            import csv
            from io import StringIO
            csv_text = "\n".join(lines[header_index:])
            reader = csv.reader(StringIO(csv_text), delimiter=delimiter)

            valid_rows = []
            for row in reader:
                if len(row) >= 3 and any(cell.strip() for cell in row):
                    valid_rows.append(row)

            if not valid_rows: continue

            header = valid_rows[0]
            data = valid_rows[1:]

            max_cols = len(header)
            clean_header = header[:max_cols]

            seen = set()
            unique_header = []
            for col in clean_header:
                col_name = col.strip()
                if col_name in seen:
                    col_name = f"{col_name}_{len(seen)}"
                seen.add(col_name)
                unique_header.append(col_name)

            clean_data = []
            for row in data:
                if len(row) < max_cols:
                    row += [''] * (max_cols - len(row))
                clean_data.append(row[:max_cols])

            df = pd.DataFrame(clean_data, columns=unique_header)
            df = df.fillna("")
            return df

        except Exception:
            continue

    raise ValueError(f"文件 {filename} 解析失败：无法识别格式或内容为空。")


def extract_global_dict(frames):
    global_dict = {}
    for df in frames:
        if "交易分类" in df.columns:
            target_col = next((c for c in df.columns if "交易对方" in c or "对方" in c or "商户名称" in c), None)
            desc_col = next((c for c in df.columns if "商品" in c or "说明" in c or "备注" in c), None)
            for _, row in df.iterrows():
                category = str(row.get("交易分类", "")).strip()
                if not category or category == "nan": continue
                if target_col:
                    target = str(row.get(target_col, "")).strip()
                    if target and target != "nan" and len(target) >= 2:
                        global_dict[target] = category
                if desc_col:
                    desc = str(row.get(desc_col, "")).strip()
                    if desc and desc != "nan" and len(desc) >= 3:
                        global_dict[desc] = category
    return global_dict


def process_data(filepath, filename, combo_rules, expand_dict, global_dict):
    df = parse_file_to_df(filepath, filename)

    if "交易分类" in df.columns or "收/付款方式" in df.columns:
        source_type = "支付宝"
    elif "微信" in str(df.columns) or "支付方式" in df.columns:
        source_type = "微信"
    else:
        source_type = "其他"

    time_col = next((c for c in df.columns if "交易时间" in c or "交易日期" in c), None)
    if not time_col: raise ValueError("找不到时间列！")

    df["时间"] = pd.to_datetime(df[time_col], errors="coerce")
    df["日期"] = df["时间"].dt.strftime('%Y-%m-%d')

    amount_col = next((col for col in df.columns if "金额" in str(col)), None)
    if not amount_col: raise ValueError("找不到金额列！")
    df["金额"] = df[amount_col].apply(clean_amount)

    # ================= 核心修改：兼容“不计收支” =================
    if "收/支" in df.columns:
        df["收支"] = df["收/支"].astype(str).str.strip()
    elif "收支" in df.columns:
        df["收支"] = df["收支"].astype(str).str.strip()
    elif "交易类型" in df.columns:
        df["收支"] = df["交易类型"].apply(
            lambda x: "收入" if any(k in str(x) for k in ["退款", "收入", "转入"]) else "支出")
    else:
        df["收支"] = "支出"

        # 支付宝特有的“不计收支”，我们将其处理为“不计收支”，并在后面特殊过滤
    df.loc[df["收支"].str.contains("不计收支", na=False), "收支"] = "不计收支"
    # ==========================================================

    if "交易状态" in df.columns:
        df = df[~df["交易状态"].astype(str).str.contains("失败", na=False)]

    target_col = next((col for col in df.columns if "交易对方" in col or "对方" in col or "商户名称" in col), None)
    desc_col = next((col for col in df.columns if "商品" in col or "说明" in col or "备注" in col), None)

    if source_type == "支付宝":
        df["原分类"] = df.get("交易分类", "其他").fillna("其他").astype(str).str.strip()
    else:
        df["原分类"] = df.get("交易类型", df.get("交易分类", "其他")).fillna("其他").astype(str).str.strip()

    categories = []
    for _, row in df.iterrows():
        target_name = str(row.get(target_col, "")).strip() if target_col else ""
        desc_text = str(row.get(desc_col, "")).strip() if desc_col else ""
        combined_text = target_name + desc_text
        original_cat = row["原分类"]
        current_shouzhi = row["收支"]

        # 【核心1】：如果是“不计收支”，先判断是不是内部流转，打上特定标签
        if current_shouzhi == "不计收支":
            if any(k in original_cat or k in combined_text for k in ["投资", "理财", "还款", "提现", "余额宝", "转账"]):
                categories.append("内部资金流转(不计入总额)")
            else:
                categories.append("不计收支(其他)")
            continue

        matched = False
        for rule in combo_rules:
            if all(t in combined_text for t in rule["targets"]):
                categories.append(rule["category"])
                matched = True
                break
        if matched: continue

        for keyword, category in expand_dict.items():
            if keyword in combined_text:
                categories.append(category)
                matched = True
                break
        if matched: continue

        if target_name in global_dict:
            categories.append(global_dict[target_name])
            continue
        if desc_text in global_dict:
            categories.append(global_dict[desc_text])
            continue

        for line in DEFAULT_FALLBACK_RULES.strip().split('\n'):
            if "=" in line:
                kw, cat = line.split("=")
                if kw.strip() in combined_text:
                    categories.append(cat.strip())
                    matched = True
                    break
        if matched: continue

        categories.append(original_cat)

    df["消费类型"] = categories
    df["来源"] = source_type
    df["交易对方"] = df[target_col].astype(str) if target_col else "未知"

    order_col = next((col for col in df.columns if "单号" in col or "流水号" in col or "订单号" in col), None)
    df["订单号"] = df[order_col].astype(str) if order_col else ""

    df = df.reset_index(drop=True)
    return df[["时间", "日期", "收支", "消费类型", "金额", "来源", "交易对方", "订单号"]]


def find_cancel_out_records(df):
    # 只在“支出”和“收入”里找对冲，排除不计收支
    expenses = df[df["收支"] == "支出"].copy()
    incomes = df[df["收支"] == "收入"].copy()
    rows_to_drop = set()
    cancel_records = []

    for idx, exp in expenses.iterrows():
        if idx in rows_to_drop: continue
        match_inc = incomes[(incomes["金额"] == exp["金额"]) &
                            (incomes["时间"] >= exp["时间"]) &
                            (incomes["时间"] <= exp["时间"] + pd.Timedelta(days=CANCEL_OUT_DAYS))]
        if not match_inc.empty:
            inc_idx = match_inc.index[0]
            inc_row = match_inc.iloc[0]
            rows_to_drop.add(idx)
            rows_to_drop.add(inc_idx)
            incomes = incomes.drop(inc_idx)

            cancel_records.append({
                "expense_idx": int(idx),
                "income_idx": int(inc_idx),
                "支出日期": exp["日期"],
                "支出金额": exp["金额"],
                "支出对方": exp["交易对方"],
                "收入日期": inc_row["日期"],
                "收入金额": inc_row["金额"],
                "收入对方": inc_row["交易对方"]
            })
    return cancel_records


# ===== 以下为 UI 与接口 =====
def auto_adjust_column_width(writer):
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for col_idx, column in enumerate(worksheet.columns, 1):
            max_length = 0
            column_letter = get_column_letter(col_idx)
            for cell in column:
                if cell.value:
                    length = sum(2 if ord(c) > 127 else 1 for c in str(cell.value))
                    max_length = max(max_length, length)
            worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 30)


def format_excel(writer):
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    month_total_font = Font(bold=True, color="9C0006")
    month_total_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for cell in worksheet[1]:
            cell.font = header_font;
            cell.fill = header_fill;
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in worksheet.iter_rows(min_row=2):
            is_total = "合计" in str(row[0].value) if row[0].value else False
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(cell.value, (int, float)): cell.number_format = '#,##0.00'
                if is_total: cell.font = month_total_font; cell.fill = month_total_fill


@app.get("/", response_class=HTMLResponse)
async def index():
    return """
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>智能记账助手 V9.1</title>
        <style>
            * { box-sizing: border-box; }
            body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #eef2f5; margin: 0; padding: 20px; display: flex; flex-direction: column; align-items: center; }
            .container { background: white; padding: 30px 40px; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.08); width: 100%; max-width: 900px; }
            h1 { color: #2c3e50; margin-top: 0; margin-bottom: 10px; text-align: center; font-size: 24px; }
            .subtitle { color: #7f8c8d; font-size: 14px; margin-bottom: 30px; text-align: center; }

            .grid-container { display: flex; gap: 20px; margin-bottom: 20px; }
            @media (max-width: 650px) { .grid-container { flex-direction: column; } }
            .card { flex: 1; background: #f8fafc; border: 2px dashed #cbd5e1; padding: 20px; border-radius: 12px; text-align: left; display: flex; flex-direction: column; }
            .card h3 { margin-top: 0; color: #1e293b; font-size: 16px; border-bottom: 1px solid #e2e8f0; padding-bottom: 10px; margin-bottom: 15px; }
            .rule-input { width: 100%; flex-grow: 1; min-height: 120px; padding: 10px; border: 1px solid #cbd5e1; border-radius: 6px; margin-bottom: 10px; font-size: 13px; }

            .submit-btn { background-color: #3b82f6; color: white; border: none; padding: 16px 24px; border-radius: 8px; cursor: pointer; font-size: 18px; font-weight: bold; width: 100%; transition: 0.3s; }
            .submit-btn:disabled { background-color: #94a3b8; cursor: not-allowed; }

            #status { margin-top: 15px; font-size: 15px; font-weight: bold; text-align: center; min-height: 22px; }

            #extraActions { margin-top: 20px; display: none; text-align: center; }
            .btn-warning { background-color: #f59e0b; color: white; border: none; padding: 12px 20px; border-radius: 8px; cursor: pointer; font-size: 15px; font-weight: bold; transition: 0.3s; box-shadow: 0 4px 6px rgba(245,158,11,0.2); }
            .btn-warning:hover { background-color: #d97706; }
            .btn-success { background-color: #10b981; color: white; border: none; padding: 12px 20px; border-radius: 8px; cursor: pointer; font-size: 15px; font-weight: bold; margin-left: 10px; box-shadow: 0 4px 6px rgba(16,185,129,0.2); }

            .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); justify-content: center; align-items: center; z-index: 1000; }
            .modal-content { background: white; padding: 25px; border-radius: 12px; width: 95%; max-width: 850px; max-height: 90vh; overflow-y: auto; }
            .modal-title { color: #d97706; font-size: 18px; font-weight: bold; margin-bottom: 10px; }
            .modal-desc { font-size: 14px; color: #475569; margin-bottom: 15px; }
            .table-container { max-height: 45vh; overflow-y: auto; border: 1px solid #cbd5e1; border-radius: 6px; margin-bottom: 20px; }
            table { width: 100%; border-collapse: collapse; font-size: 13px; }
            th, td { padding: 10px 12px; border-bottom: 1px solid #e2e8f0; text-align: left; }
            th { background: #f1f5f9; position: sticky; top: 0; }
            input[type="text"] { padding: 6px; border: 1px solid #cbd5e1; border-radius: 4px; width: 120px; }
            .modal-actions { display: flex; gap: 15px; justify-content: flex-end; }
            .btn-cancel { background: #f1f5f9; color: #475569; border: 1px solid #cbd5e1; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: bold; }
            .btn-confirm { background: #10b981; color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: bold; }
            .multi-select-tip { background-color: #fef08a; color: #854d0e; padding: 6px 10px; border-radius: 6px; font-size: 13px; font-weight: bold; margin-bottom: 10px; display: inline-block; }
            input[type="checkbox"] { width: 18px; height: 18px; cursor: pointer; accent-color: #3b82f6; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 智能记账助手 (V9.1)</h1>
            <div class="subtitle">完美过滤内部资金流转，真假 Excel 通吃</div>

            <div class="grid-container">
                <div class="card" style="flex: 0.8;">
                    <h3>📁 1. 上传账单</h3>
                    <div class="multi-select-tip">💡 按住 Shift 或 Ctrl 键可多选！</div>
                    <input type="file" id="billFiles" class="file-input" accept=".csv, .xlsx, .xls" multiple>
                </div>
                <div class="card" style="flex: 1.2;">
                    <h3>🧠 2. 强制覆盖规则 (选填)</h3>
                    <textarea id="ruleText" class="rule-input" placeholder="例如：华北水利水电大学+校园卡=校园卡充值"></textarea>
                </div>
            </div>

            <button class="submit-btn" id="submitBtn" onclick="analyzeFiles()">第一步：分析账单与代付对冲</button>

            <div id="extraActions">
                <button id="vagueBtn" class="btn-warning" onclick="openVagueModal()" style="display:none;">🔍 发现未明确分类，去手动归类</button>
                <button id="generateBtn" class="btn-success" onclick="finalGenerate()">✅ 确认无误，生成 Excel</button>
            </div>

            <div id="status"></div>
        </div>

        <div id="reviewModal" class="modal-overlay">
            <div class="modal-content">
                <div class="modal-title">⚠️ 发现可能的代付/垫付记录</div>
                <div class="modal-desc">请勾选确认要抵消的记录：</div>
                <div class="table-container">
                    <table id="reviewTable">
                        <thead>
                            <tr>
                                <th style="width: 40px;"><input type="checkbox" id="selectAll" onchange="toggleAll(this)" checked></th>
                                <th>支出日期/金额</th>
                                <th>支出对方</th>
                                <th>收入日期/金额</th>
                                <th>收入对方</th>
                            </tr>
                        </thead>
                        <tbody id="reviewTableBody"></tbody>
                    </table>
                </div>
                <div class="modal-actions">
                    <button class="btn-cancel" onclick="cancelAllCancel()">❌ 全部不抵消</button>
                    <button class="btn-confirm" onclick="confirmSelectedCancel()">✅ 确认勾选项</button>
                </div>
            </div>
        </div>

        <div id="vagueModal" class="modal-overlay">
            <div class="modal-content">
                <div class="modal-title">✍️ 手动指定账单分类</div>
                <div class="modal-desc">这些账单程序无法猜透，请输入你想要的分类名称：</div>
                <div class="table-container">
                    <table id="vagueTable">
                        <thead>
                            <tr>
                                <th>时间</th>
                                <th>对方</th>
                                <th>金额</th>
                                <th>原分类</th>
                                <th>手动归类为：</th>
                            </tr>
                        </thead>
                        <tbody id="vagueTableBody"></tbody>
                    </table>
                </div>
                <div class="modal-actions">
                    <button class="btn-cancel" onclick="document.getElementById('vagueModal').style.display='none'">取消</button>
                    <button class="btn-confirm" onclick="saveVagueChanges()">保存归类</button>
                </div>
            </div>
        </div>

        <script>
            let tempData = null;
            let selectedCancelIndexes = []; 
            let currentVagueRecords = []; 

            async function analyzeFiles() {
                const billFiles = document.getElementById('billFiles').files;
                const ruleText = document.getElementById('ruleText').value;
                const statusDiv = document.getElementById('status');
                const submitBtn = document.getElementById('submitBtn');

                if (billFiles.length === 0) {
                    statusDiv.innerText = "⚠️ 请先选择账单文件！";
                    statusDiv.style.color = "#dc2626";
                    return;
                }

                const formData = new FormData();
                for (let i = 0; i < billFiles.length; i++) { formData.append("bill_files", billFiles[i]); }
                if (ruleText.trim()) { formData.append("rule_text", ruleText); }

                statusDiv.innerText = "正在分析账单...";
                statusDiv.style.color = "#1e293b";
                submitBtn.disabled = true;

                try {
                    const response = await fetch('/analyze', { method: 'POST', body: formData });
                    if (!response.ok) throw new Error(await response.text());

                    tempData = await response.json();

                    if (tempData.cancel_records.length === 0) {
                        checkVagueRecords();
                    } else {
                        showReviewModal(tempData.cancel_records);
                        statusDiv.innerText = "请核对抵消记录。";
                    }
                } catch (error) {
                    statusDiv.innerText = "❌ " + error.message;
                    statusDiv.style.color = "#dc2626";
                    submitBtn.disabled = false;
                }
            }

            function showReviewModal(records) {
                const tbody = document.getElementById('reviewTableBody');
                tbody.innerHTML = "";
                records.forEach((r, i) => {
                    tbody.innerHTML += `<tr>
                        <td><input type="checkbox" class="cancel-checkbox" data-idx="${i}" checked></td>
                        <td style="color:#dc2626">-${r['支出金额']}元<br><small>${r['支出日期']}</small></td>
                        <td>${r['支出对方']}</td>
                        <td style="color:#16a34a">+${r['收入金额']}元<br><small>${r['收入日期']}</small></td>
                        <td>${r['收入对方']}</td>
                    </tr>`;
                });
                document.getElementById('selectAll').checked = true;
                document.getElementById('reviewModal').style.display = 'flex';
            }

            function toggleAll(source) {
                document.querySelectorAll('.cancel-checkbox').forEach(cb => cb.checked = source.checked);
            }

            function cancelAllCancel() {
                document.getElementById('reviewModal').style.display = 'none';
                selectedCancelIndexes = []; 
                checkVagueRecords();
            }

            function confirmSelectedCancel() {
                const checkboxes = document.querySelectorAll('.cancel-checkbox');
                selectedCancelIndexes = [];
                checkboxes.forEach(cb => {
                    if (cb.checked) selectedCancelIndexes.push(parseInt(cb.getAttribute('data-idx')));
                });
                document.getElementById('reviewModal').style.display = 'none';
                checkVagueRecords();
            }

            async function checkVagueRecords() {
                const statusDiv = document.getElementById('status');
                statusDiv.innerText = "正在检查模糊分类...";

                try {
                    const response = await fetch('/check_vague', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ session_id: tempData.session_id })
                    });

                    const data = await response.json();
                    currentVagueRecords = data.vague_records;

                    document.getElementById('extraActions').style.display = 'block';

                    if (currentVagueRecords.length > 0) {
                        document.getElementById('vagueBtn').style.display = 'inline-block';
                        document.getElementById('vagueBtn').innerText = `🔍 发现 ${currentVagueRecords.length} 笔未明确分类，去手动归类`;
                        statusDiv.innerText = "请点击黄色按钮处理未分类账单，或直接点击绿色按钮生成。";
                    } else {
                        document.getElementById('vagueBtn').style.display = 'none';
                        statusDiv.innerText = "未发现模糊分类，可以直接生成！";
                    }
                    document.getElementById('status').style.color = "#1e293b";

                } catch (e) {
                    statusDiv.innerText = "❌ 检查失败：" + e.message;
                }
            }

            function openVagueModal() {
                const tbody = document.getElementById('vagueTableBody');
                tbody.innerHTML = "";
                currentVagueRecords.forEach((r, i) => {
                    tbody.innerHTML += `<tr>
                        <td>${r['日期']}</td>
                        <td>${r['交易对方']}</td>
                        <td style="color:#dc2626">${r['金额']}元</td>
                        <td><span style="background:#e2e8f0; padding:2px 6px; border-radius:4px;">${r['消费类型']}</span></td>
                        <td><input type="text" class="vague-input" data-idx="${r['df_idx']}" placeholder="输入新分类"></td>
                    </tr>`;
                });
                document.getElementById('vagueModal').style.display = 'flex';
            }

            async function saveVagueChanges() {
                const inputs = document.querySelectorAll('.vague-input');
                const updates = [];
                inputs.forEach(inp => {
                    if (inp.value.trim() !== "") {
                        updates.push({
                            df_idx: parseInt(inp.getAttribute('data-idx')),
                            new_category: inp.value.trim()
                        });
                    }
                });

                if (updates.length === 0) {
                    document.getElementById('vagueModal').style.display = 'none';
                    return;
                }

                try {
                    const response = await fetch('/update_vague', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ 
                            session_id: tempData.session_id,
                            updates: updates 
                        })
                    });
                    if (!response.ok) throw new Error(await response.text());

                    document.getElementById('vagueModal').style.display = 'none';
                    document.getElementById('vagueBtn').style.display = 'none';
                    document.getElementById('status').innerText = "✅ 归类已保存！可以生成 Excel 了。";
                    document.getElementById('status').style.color = "#16a34a";
                } catch (e) {
                    alert("保存失败：" + e.message);
                }
            }

            async function finalGenerate() {
                const statusDiv = document.getElementById('status');
                const submitBtn = document.getElementById('submitBtn');
                document.getElementById('extraActions').style.display = 'none';
                statusDiv.innerText = "正在生成最终 Excel...";

                try {
                    const response = await fetch('/generate', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            session_id: tempData.session_id,
                            cancel_indexes: selectedCancelIndexes
                        })
                    });

                    if (!response.ok) throw new Error(await response.text());

                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url; a.download = "我的消费记账表.xlsx";
                    document.body.appendChild(a); a.click(); a.remove();

                    statusDiv.innerText = "✅ 生成成功！已开始下载。";
                    statusDiv.style.color = "#16a34a";
                    submitBtn.disabled = false;
                } catch (error) {
                    statusDiv.innerText = "❌ " + error.message;
                    statusDiv.style.color = "#dc2626";
                    submitBtn.disabled = false;
                }
            }
        </script>
    </body>
    </html>
    """


@app.post("/analyze")
async def analyze_files(bill_files: list[UploadFile] = File(...), rule_text: str = Form(None)):
    session_id = str(uuid.uuid4().hex)
    temp_dir = f"temp_{session_id}"
    os.makedirs(temp_dir, exist_ok=True)

    combo_rules, expand_dict = parse_user_rules(rule_text)
    frames = []
    for file in bill_files:
        filepath = os.path.join(temp_dir, file.filename)
        with open(filepath, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        try:
            frames.append(parse_file_to_df(filepath, file.filename))
        except Exception as e:
            raise ValueError(f"读取文件 {file.filename} 失败: {str(e)}")

    global_dict = extract_global_dict(frames)

    processed_frames = []
    for file in bill_files:
        filepath = os.path.join(temp_dir, file.filename)
        try:
            processed_frames.append(process_data(filepath, file.filename, combo_rules, expand_dict, global_dict))
        except Exception as e:
            raise ValueError(f"处理数据 {file.filename} 失败: {str(e)}")

    all_df = pd.concat(processed_frames, ignore_index=True).dropna(subset=["日期", "金额"])

    # 不要在这里剔除“不计收支”，因为用户可能在“手动归类的提示”里看到它
    # 但是为了去重，可以保留
    all_df = all_df.drop_duplicates(subset=["时间", "金额", "交易对方", "订单号"])
    all_df = all_df.reset_index(drop=True)

    all_df.to_pickle(os.path.join(temp_dir, "data.pkl"))
    cancel_records = find_cancel_out_records(all_df)

    return {"session_id": session_id, "cancel_records": cancel_records}


@app.post("/check_vague")
async def check_vague(payload: dict = Body(...)):
    session_id = payload.get("session_id")
    temp_dir = f"temp_{session_id}"
    pkl_path = os.path.join(temp_dir, "data.pkl")

    if not os.path.exists(pkl_path):
        return {"vague_records": []}

    all_df = pd.read_pickle(pkl_path)
    # 只针对真正的支出寻找模糊分类，排除内部流转
    expense_df = all_df[(all_df["收支"] == "支出") & (~all_df["消费类型"].str.contains("内部资金", na=False))]
    vague_mask = expense_df["消费类型"].isin(VAGUE_CATEGORIES)
    vague_df = expense_df[vague_mask]

    records = []
    for idx, row in vague_df.iterrows():
        records.append({
            "df_idx": int(idx),
            "日期": row["日期"],
            "交易对方": row["交易对方"],
            "金额": row["金额"],
            "消费类型": row["消费类型"]
        })

    return {"vague_records": records}


@app.post("/update_vague")
async def update_vague(payload: dict = Body(...)):
    session_id = payload.get("session_id")
    updates = payload.get("updates", [])

    temp_dir = f"temp_{session_id}"
    pkl_path = os.path.join(temp_dir, "data.pkl")

    if os.path.exists(pkl_path):
        all_df = pd.read_pickle(pkl_path)
        for item in updates:
            idx = item["df_idx"]
            new_cat = item["new_category"]
            if idx in all_df.index:
                all_df.at[idx, "消费类型"] = new_cat
        all_df.to_pickle(pkl_path)

    return {"status": "success"}


@app.post("/generate")
async def generate_excel(payload: dict = Body(...)):
    session_id = payload.get("session_id")
    cancel_indexes = payload.get("cancel_indexes", [])

    temp_dir = f"temp_{session_id}"
    pkl_path = os.path.join(temp_dir, "data.pkl")

    if not os.path.exists(pkl_path):
        return HTMLResponse(content="会话过期或无效，请重新上传。", status_code=400)

    all_df = pd.read_pickle(pkl_path)
    all_cancel_records = find_cancel_out_records(all_df)

    rows_to_drop = set()
    selected_cancel_records = []

    for i in cancel_indexes:
        if i < len(all_cancel_records):
            record = all_cancel_records[i]
            selected_cancel_records.append(record)
            rows_to_drop.add(record["expense_idx"])
            rows_to_drop.add(record["income_idx"])

    # 剔除被抵消的行
    final_df = all_df.drop(index=list(rows_to_drop), errors='ignore')

    # 【核心2】：最终生成表格时，剔除所有的“收入”和“不计收支”，只留“支出”
    final_df = final_df[final_df["收支"] == "支出"]

    if final_df.empty:
        shutil.rmtree(temp_dir, True)
        return HTMLResponse(content="没有找到有效的支出记录。", status_code=400)

    final_df = final_df.drop_duplicates(subset=["时间", "金额", "交易对方", "订单号"])
    final_df["月份"] = final_df["日期"].str[:7]

    # 如果用户强行把“不计收支”手动改成了“餐饮”，它会算进“支出”里。
    # 如果是内部流转，它的标签是“内部资金流转(不计入总额)”，我们也不把它放入 daily_pivot 统计（但保留在明细里）
    stat_df = final_df[~final_df["消费类型"].str.contains("内部资金", na=False)]

    daily_pivot = stat_df.pivot_table(index="日期", columns="消费类型", values="金额", aggfunc="sum", fill_value=0)
    daily_pivot["当日合计"] = daily_pivot.sum(axis=1)
    daily_pivot = daily_pivot.sort_index()

    merged_records = []
    for month, month_data in daily_pivot.groupby(daily_pivot.index.str[:7]):
        for date_str, row in month_data.iterrows():
            record = row.to_dict();
            record["日期/月份"] = date_str;
            merged_records.append(record)
        month_total_record = month_data.sum().to_dict();
        month_total_record["日期/月份"] = f"【{month} 月总计】";
        merged_records.append(month_total_record)

    merged_df = pd.DataFrame(merged_records)
    if not merged_df.empty:
        cols = ["日期/月份"] + [c for c in merged_df.columns if c != "日期/月份"]
        merged_df = merged_df[cols]

    output_path = os.path.join(temp_dir, "result.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if not merged_df.empty:
            merged_df.to_excel(writer, sheet_name="每日与每月完整账本", index=False)
        if selected_cancel_records:
            clean_records = []
            for r in selected_cancel_records:
                clean_records.append({
                    "支出日期": r["支出日期"],
                    "支出金额": r["支出金额"],
                    "支出对方": r["支出对方"],
                    "收入日期": r["收入日期"],
                    "收入金额": r["收入金额"],
                    "收入对方": r["收入对方"]
                })
            df_cancel = pd.DataFrame(clean_records)
            df_cancel.to_excel(writer, sheet_name="已确认的代付对冲明细", index=False)

        # 最终明细里，我们会展示所有流水（包括内部流转），方便用户审计
        final_df[["日期", "月份", "消费类型", "金额", "来源"]].to_excel(writer, sheet_name="真实支出明细", index=False)
        auto_adjust_column_width(writer)
        format_excel(writer)

    with open(output_path, "rb") as f:
        file_data = f.read()

    shutil.rmtree(temp_dir, True)

    return HTMLResponse(
        content=file_data,
        headers={
            "Content-Disposition": "attachment; filename=my_bill_report.xlsx",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
    )


if __name__ == "__main__":
    import uvicorn

    print("🚀 网页版记账助手已启动！")
    print("👉 请在浏览器中打开：http://127.0.0.1:8000")
    uvicorn.run(app, host="127.0.0.1", port=8000)
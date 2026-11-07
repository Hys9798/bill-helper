import pandas as pd
import os
import shutil
import uuid
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
衣=服饰鞋包
裤=服饰鞋包
鞋=服饰鞋包
淘宝=网购
"""


def clean_amount(val):
    if pd.isna(val): return 0.0
    s = str(val).replace("¥", "").replace("`", "").replace(",", "").replace('"', '').strip()
    # 处理可能带的正负号
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


def smart_categorize(row, target_col, desc_col, combo_rules, expand_dict, fallback_rules):
    target_name = str(row.get(target_col, "")).strip() if target_col else ""
    desc_text = str(row.get(desc_col, "")).strip() if desc_col else ""
    combined_text = target_name + desc_text

    for rule in combo_rules:
        if all(t in combined_text for t in rule["targets"]):
            return rule["category"]
    for keyword, category in expand_dict.items():
        if keyword in combined_text:
            return category
    for line in fallback_rules.strip().split('\n'):
        if "=" in line:
            kw, cat = line.split("=")
            if kw.strip() in combined_text:
                return cat.strip()
    return row["原分类"]


def process_uploaded_file(filepath, filename, combo_rules, expand_dict):
    ext = Path(filepath).suffix.lower()
    df = None

    if ext in ['.xlsx', '.xls']:
        df_raw = pd.read_excel(filepath, header=None)
        header_row = -1
        for i in range(min(50, len(df_raw))):
            # 兼容云闪付的“交易日期”或微信支付宝的“交易时间”
            row_str = str(df_raw.iloc[i].values)
            if "交易时间" in row_str or "交易日期" in row_str:
                header_row = i
                break
        if header_row == -1: raise ValueError(f"文件 {filename} 中找不到【交易时间】或【交易日期】表头！")
        df = pd.read_excel(filepath, header=header_row)
    else:
        success = False
        for encoding in ["utf-8-sig", "gbk", "gb18030", "utf-8"]:
            try:
                with open(filepath, 'r', encoding=encoding) as f:
                    lines = f.readlines()
                header_index = -1
                for i, line in enumerate(lines):
                    if "交易时间" in line or "交易日期" in line:
                        header_index = i
                        break
                if header_index != -1:
                    import csv
                    from io import StringIO
                    csv_text = "".join(lines[header_index:])
                    reader = csv.reader(StringIO(csv_text))
                    valid_lines = [row for row in reader if len(row) > 0]
                    if valid_lines:
                        df = pd.DataFrame(valid_lines[1:], columns=valid_lines[0])
                        df = df.fillna("")
                        success = True
                        break
            except Exception:
                continue
        if not success: raise ValueError(f"文件 {filename} 无法识别！")

    df.columns = [str(col).strip().replace('\n', '') for col in df.columns]

    # 识别平台
    if "交易分类" in df.columns or "收/付款方式" in df.columns:
        source_type = "支付宝"
    elif "微信" in str(df.columns):
        source_type = "微信"
    elif "交易流水号" in df.columns or "商户名称" in df.columns:
        source_type = "云闪付"
    else:
        source_type = "其他"

    # 智能识别时间列
    time_col = next((c for c in df.columns if "交易时间" in c or "交易日期" in c), None)
    if not time_col: raise ValueError("找不到时间列！")
    df["时间"] = pd.to_datetime(df[time_col], errors="coerce")
    df["日期"] = df["时间"].dt.strftime('%Y-%m-%d')

    # 智能识别金额列
    amount_col = next((col for col in df.columns if "金额" in str(col)), None)
    if not amount_col: raise ValueError(f"找不到金额列！")
    df["金额"] = df[amount_col].apply(clean_amount)

    # 智能判断收支
    if "收/支" in df.columns:
        df["收支"] = df["收/支"].astype(str).str.strip()
    elif "收支" in df.columns:
        df["收支"] = df["收支"].astype(str).str.strip()
    elif "交易类型" in df.columns:
        # 云闪付的逻辑：如果不是退款/收入，就算支出
        df["收支"] = df["交易类型"].apply(
            lambda x: "收入" if any(k in str(x) for k in ["退款", "收入", "转入"]) else "支出")
    else:
        df["收支"] = "支出"

        # 过滤掉失败交易
    if "交易状态" in df.columns:
        df = df[~df["交易状态"].astype(str).str.contains("失败", na=False)]

    # 智能识别分类列
    if source_type == "支付宝":
        df["原分类"] = df.get("交易分类", df.get("交易类型", "其他")).fillna("其他").astype(str).str.strip()
    else:
        df["原分类"] = df.get("交易类型", df.get("交易分类", "其他")).fillna("其他").astype(str).str.strip()

    # 智能识别对方名称
    target_col = next((col for col in df.columns if "交易对方" in col or "对方" in col or "商户名称" in col), None)
    desc_col = next((col for col in df.columns if "商品" in col or "说明" in col or "备注" in col), None)

    df["消费类型"] = df.apply(
        lambda row: smart_categorize(row, target_col, desc_col, combo_rules, expand_dict, DEFAULT_FALLBACK_RULES),
        axis=1)

    df["来源"] = source_type
    df["交易对方"] = df[target_col].astype(str) if target_col else "未知"

    df = df.reset_index(drop=True)
    return df[["时间", "日期", "收支", "消费类型", "金额", "来源", "交易对方"]]


def find_cancel_out_records(df):
    expenses = df[df["收支"].str.contains("支出", na=False)].copy()
    incomes = df[df["收支"].str.contains("收入", na=False)].copy()
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
        <title>智能记账助手 V6.2</title>
        <style>
            * { box-sizing: border-box; }
            body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #eef2f5; margin: 0; padding: 20px; display: flex; flex-direction: column; align-items: center; }
            .container { background: white; padding: 30px 40px; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.08); width: 100%; max-width: 900px; }
            h1 { color: #2c3e50; margin-top: 0; margin-bottom: 10px; text-align: center; font-size: 24px; }
            .subtitle { color: #7f8c8d; font-size: 14px; margin-bottom: 30px; text-align: center; }

            .grid-container { display: flex; gap: 20px; margin-bottom: 30px; }
            @media (max-width: 650px) { .grid-container { flex-direction: column; } }
            .card { flex: 1; background: #f8fafc; border: 2px dashed #cbd5e1; padding: 20px; border-radius: 12px; text-align: left; display: flex; flex-direction: column; }
            .card h3 { margin-top: 0; color: #1e293b; font-size: 16px; border-bottom: 1px solid #e2e8f0; padding-bottom: 10px; margin-bottom: 15px; }
            .rule-input { width: 100%; flex-grow: 1; min-height: 120px; padding: 10px; border: 1px solid #cbd5e1; border-radius: 6px; margin-bottom: 10px; font-size: 13px; }

            .submit-btn { background-color: #3b82f6; color: white; border: none; padding: 16px 24px; border-radius: 8px; cursor: pointer; font-size: 18px; font-weight: bold; width: 100%; }
            .submit-btn:disabled { background-color: #94a3b8; cursor: not-allowed; }

            #status { margin-top: 15px; font-size: 15px; font-weight: bold; text-align: center; min-height: 22px; }

            .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); justify-content: center; align-items: center; z-index: 1000; }
            .modal-content { background: white; padding: 25px; border-radius: 12px; width: 95%; max-width: 850px; max-height: 90vh; overflow-y: auto; }
            .modal-title { color: #d97706; font-size: 18px; font-weight: bold; margin-bottom: 10px; display: flex; align-items: center; gap: 8px; }
            .modal-desc { font-size: 14px; color: #475569; margin-bottom: 15px; }
            .table-container { max-height: 45vh; overflow-y: auto; border: 1px solid #cbd5e1; border-radius: 6px; margin-bottom: 20px; }
            table { width: 100%; border-collapse: collapse; font-size: 13px; }
            th, td { padding: 10px 12px; border-bottom: 1px solid #e2e8f0; text-align: left; }
            th { background: #f1f5f9; position: sticky; top: 0; }
            tr:hover { background-color: #f8fafc; }
            .modal-actions { display: flex; gap: 15px; justify-content: flex-end; }
            .btn-cancel { background: #f1f5f9; color: #475569; border: 1px solid #cbd5e1; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: bold; }
            .btn-confirm { background: #10b981; color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: bold; }
            .multi-select-tip { background-color: #fef08a; color: #854d0e; padding: 6px 10px; border-radius: 6px; font-size: 13px; font-weight: bold; margin-bottom: 10px; display: inline-block; }
            input[type="checkbox"] { width: 18px; height: 18px; cursor: pointer; accent-color: #3b82f6; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 智能记账助手 (V6.2 融合版)</h1>
            <div class="subtitle">完美兼容微信、支付宝与云闪付</div>

            <div class="grid-container">
                <div class="card" style="flex: 0.8;">
                    <h3>📁 1. 上传账单</h3>
                    <div class="multi-select-tip">💡 提示：按住 Shift 或 Ctrl 键可多选！</div>
                    <input type="file" id="billFiles" class="file-input" accept=".csv, .xlsx, .xls" multiple>
                </div>
                <div class="card" style="flex: 1.2;">
                    <h3>🧠 2. 强制覆盖规则 (选填)</h3>
                    <textarea id="ruleText" class="rule-input" placeholder="例如：华北水利水电大学+校园卡=校园卡充值&#10;五金店=餐饮"></textarea>
                </div>
            </div>

            <button class="submit-btn" id="submitBtn" onclick="analyzeFiles()">第一步：分析账单与代付对冲</button>
            <div id="status"></div>
        </div>

        <div id="reviewModal" class="modal-overlay">
            <div class="modal-content">
                <div class="modal-title">⚠️ 发现可能的代付/垫付记录</div>
                <div class="modal-desc">程序发现以下金额匹配且时间相近的收支记录。<b>请勾选确认要抵消的记录：</b></div>
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
                    <button class="btn-cancel" onclick="cancelAll()">❌ 全部不抵消</button>
                    <button class="btn-confirm" onclick="confirmSelected()">✅ 确认勾选项并生成</button>
                </div>
            </div>
        </div>

        <script>
            let tempData = null;

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

                statusDiv.innerText = "正在分析账单并寻找代付记录...";
                statusDiv.style.color = "#1e293b";
                submitBtn.disabled = true;

                try {
                    const response = await fetch('/analyze', { method: 'POST', body: formData });
                    if (!response.ok) throw new Error(await response.text());

                    tempData = await response.json();

                    if (tempData.cancel_records.length === 0) {
                        statusDiv.innerText = "未发现代付记录，正在直接生成...";
                        generateFinalExcel([]);
                    } else {
                        showModal(tempData.cancel_records);
                        statusDiv.innerText = "请核对抵消记录后点击生成。";
                    }
                } catch (error) {
                    statusDiv.innerText = "❌ " + error.message;
                    statusDiv.style.color = "#dc2626";
                    submitBtn.disabled = false;
                }
            }

            function showModal(records) {
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
                const checkboxes = document.querySelectorAll('.cancel-checkbox');
                checkboxes.forEach(cb => cb.checked = source.checked);
            }

            function cancelAll() {
                document.getElementById('reviewModal').style.display = 'none';
                document.getElementById('status').innerText = "已取消所有抵消，正在按纯流水生成...";
                generateFinalExcel([]); 
            }

            function confirmSelected() {
                const checkboxes = document.querySelectorAll('.cancel-checkbox');
                const selectedIndexes = [];
                checkboxes.forEach(cb => {
                    if (cb.checked) {
                        selectedIndexes.push(parseInt(cb.getAttribute('data-idx')));
                    }
                });

                document.getElementById('reviewModal').style.display = 'none';
                if (selectedIndexes.length === 0) {
                    document.getElementById('status').innerText = "未勾选抵消记录，正在按纯流水生成...";
                } else {
                    document.getElementById('status').innerText = `已确认抵消 ${selectedIndexes.length} 笔记录，正在生成...`;
                }
                generateFinalExcel(selectedIndexes);
            }

            async function generateFinalExcel(selectedIndexes) {
                const statusDiv = document.getElementById('status');
                const submitBtn = document.getElementById('submitBtn');

                try {
                    const response = await fetch('/generate', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            session_id: tempData.session_id,
                            cancel_indexes: selectedIndexes
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
                } catch (error) {
                    statusDiv.innerText = "❌ " + error.message;
                    statusDiv.style.color = "#dc2626";
                } finally {
                    submitBtn.disabled = false;
                }
            }
        </script>
    </body>
    </html>
    """


@app.post("/analyze")
async def analyze_files(
        bill_files: list[UploadFile] = File(...),
        rule_text: str = Form(None)
):
    session_id = str(uuid.uuid4().hex)
    temp_dir = f"temp_{session_id}"
    os.makedirs(temp_dir, exist_ok=True)

    combo_rules, expand_dict = parse_user_rules(rule_text)
    frames = []
    for file in bill_files:
        filepath = os.path.join(temp_dir, file.filename)
        with open(filepath, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        frames.append(process_uploaded_file(filepath, file.filename, combo_rules, expand_dict))

    all_df = pd.concat(frames, ignore_index=True).dropna(subset=["日期", "金额"])
    all_df.to_pickle(os.path.join(temp_dir, "data.pkl"))

    cancel_records = find_cancel_out_records(all_df)

    return {
        "session_id": session_id,
        "cancel_records": cancel_records
    }


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

    selected_cancel_records = []
    rows_to_drop = set()

    for i in cancel_indexes:
        if i < len(all_cancel_records):
            record = all_cancel_records[i]
            selected_cancel_records.append(record)
            rows_to_drop.add(record["expense_idx"])
            rows_to_drop.add(record["income_idx"])

    final_df = all_df.drop(rows_to_drop)
    final_df = final_df[final_df["收支"].str.contains("支出", na=False)]

    if final_df.empty:
        shutil.rmtree(temp_dir, True)
        return HTMLResponse(content="没有找到有效的支出记录。", status_code=400)

    final_df = final_df.drop_duplicates(subset=["日期", "金额", "消费类型", "来源"])
    final_df["月份"] = final_df["日期"].str[:7]

    daily_pivot = final_df.pivot_table(index="日期", columns="消费类型", values="金额", aggfunc="sum", fill_value=0)
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
    cols = ["日期/月份"] + [c for c in merged_df.columns if c != "日期/月份"]
    merged_df = merged_df[cols]

    output_path = os.path.join(temp_dir, "result.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
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
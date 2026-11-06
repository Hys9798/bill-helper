import pandas as pd
import os
import shutil
import uuid
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, Form
from fastapi.responses import FileResponse, HTMLResponse
from pathlib import Path
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
from io import BytesIO

app = FastAPI(title="智能记账助手")

DEFAULT_MAPPING = {
    "便利店": "零食购物",
    "美团": "外卖/餐饮",
}


def clean_amount(val):
    if pd.isna(val): return 0.0
    s = str(val).replace("¥", "").replace("`", "").replace(",", "").replace('"', '').strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def process_uploaded_file(filepath, filename, custom_mapping):
    ext = Path(filepath).suffix.lower()
    df = None

    if ext in ['.xlsx', '.xls']:
        df_raw = pd.read_excel(filepath, header=None)
        header_row = -1
        for i in range(min(50, len(df_raw))):
            if any("交易时间" in str(cell) for cell in df_raw.iloc[i]):
                header_row = i
                break
        if header_row == -1: raise ValueError(f"文件 {filename} 中找不到【交易时间】表头！")
        df = pd.read_excel(filepath, header=header_row)
    else:
        for encoding in ["utf-8", "gbk", "gb18030"]:
            try:
                with open(filepath, 'r', encoding=encoding) as f:
                    lines = f.readlines()
                header_index = next((i for i, line in enumerate(lines) if "交易时间" in line and "金额" in line), -1)
                if header_index != -1:
                    df = pd.read_csv(BytesIO("".join(lines[header_index:]).encode(encoding)), on_bad_lines="skip")
                    break
            except Exception:
                continue
        if df is None: raise ValueError(f"文件 {filename} 无法识别！")

    df.columns = [str(col).strip().replace('\n', '') for col in df.columns]

    if "交易分类" in df.columns or "收/付款方式" in df.columns:
        source_type = "支付宝"
    else:
        source_type = "微信"

    df["时间"] = pd.to_datetime(df["交易时间"], errors="coerce")
    df["日期"] = df["时间"].dt.strftime('%Y-%m-%d')

    amount_col = next((col for col in df.columns if "金额" in str(col)), None)
    df["金额"] = df[amount_col].apply(clean_amount)

    if "收/支" in df.columns:
        df["收支"] = df["收/支"].astype(str).str.strip()
    elif "收支" in df.columns:
        df["收支"] = df["收支"].astype(str).str.strip()
    else:
        df["收支"] = "支出"

    if source_type == "支付宝":
        df["消费类型"] = df.get("交易分类", df.get("交易类型", "其他")).fillna("其他").astype(str).str.strip()
    else:
        df["消费类型"] = df.get("交易类型", df.get("交易分类", "其他")).fillna("其他").astype(str).str.strip()

    target_col = next((col for col in df.columns if "交易对方" in col or "对方" in col or "商户" in col), None)

    if target_col:
        def apply_custom_mapping(row):
            target_name = str(row[target_col])
            for keyword, custom_type in custom_mapping.items():
                if keyword in target_name:
                    return custom_type
            return row["消费类型"]

        df["消费类型"] = df.apply(apply_custom_mapping, axis=1)

    df["来源"] = source_type
    return df[["时间", "日期", "收支", "消费类型", "金额", "来源"]]


def cancel_out_transactions(df):
    expenses = df[df["收支"].str.contains("支出", na=False)].copy()
    incomes = df[df["收支"].str.contains("收入", na=False)].copy()
    rows_to_drop = set()
    for idx, exp in expenses.iterrows():
        if idx in rows_to_drop: continue
        match_inc = incomes[(incomes["金额"] == exp["金额"]) & (incomes["时间"] >= exp["时间"]) & (
                    incomes["时间"] <= exp["时间"] + pd.Timedelta(days=7))]
        if not match_inc.empty:
            inc_idx = match_inc.index[0]
            rows_to_drop.add(idx);
            rows_to_drop.add(inc_idx)
            incomes = incomes.drop(inc_idx)
    final_df = df.drop(rows_to_drop)
    return final_df[final_df["收支"].str.contains("支出", na=False)]


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


# ================= Web 接口 =================

@app.get("/download_template")
async def download_template(background_tasks: BackgroundTasks):
    """生成并下载分类规则模板（修复版）"""
    df = pd.DataFrame({
        "交易对方(包含的关键词)": ["萧峰", "星巴克", "滴滴"],
        "对应的消费类型": ["人情往来/代付", "咖啡饮品", "交通出行"]
    })

    # 创建一个临时文件夹来存模板
    temp_dir = f"temp_{uuid.uuid4().hex}"
    os.makedirs(temp_dir, exist_ok=True)
    temp_file_path = os.path.join(temp_dir, "分类规则模板.xlsx")

    # 写入 Excel
    with pd.ExcelWriter(temp_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="分类规则")

    # 返回文件，下载完毕后自动清理临时文件夹
    background_tasks.add_task(shutil.rmtree, temp_dir, True)

    return FileResponse(
        temp_file_path,
        filename="分类规则模板.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.get("/", response_class=HTMLResponse)
async def index():
    return """
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>智能记账助手</title>
        <style>
            * { box-sizing: border-box; }
            body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #eef2f5; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px; }
            .container { background: white; padding: 30px 40px; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.08); text-align: center; max-width: 800px; width: 100%; }
            h1 { color: #2c3e50; margin-top: 0; margin-bottom: 10px; font-size: 24px; }
            .subtitle { color: #7f8c8d; font-size: 14px; margin-bottom: 30px; }

            .grid-container { display: flex; gap: 20px; margin-bottom: 30px; }
            @media (max-width: 650px) { .grid-container { flex-direction: column; } }

            .card { flex: 1; background: #f8fafc; border: 2px dashed #cbd5e1; padding: 20px; border-radius: 12px; text-align: left; transition: border-color 0.3s; display: flex; flex-direction: column; }
            .card:hover { border-color: #3b82f6; background: #f0f9ff; }
            .card h3 { margin-top: 0; color: #1e293b; font-size: 16px; border-bottom: 1px solid #e2e8f0; padding-bottom: 10px; margin-bottom: 15px; }
            .file-input { margin: 10px 0; width: 100%; color: #475569; }

            .rule-input { width: 100%; flex-grow: 1; min-height: 100px; padding: 10px; border: 1px solid #cbd5e1; border-radius: 6px; font-family: inherit; resize: vertical; margin-bottom: 10px; font-size: 13px; }
            .rule-input:focus { outline: none; border-color: #3b82f6; box-shadow: 0 0 0 2px rgba(59,130,246,0.2); }

            .btn-link { display: inline-block; background-color: #64748b; color: white; text-decoration: none; padding: 6px 10px; border-radius: 4px; font-size: 12px; transition: 0.2s; }
            .btn-link:hover { background-color: #475569; }

            .submit-btn { background-color: #3b82f6; color: white; border: none; padding: 16px 24px; border-radius: 8px; cursor: pointer; font-size: 18px; transition: 0.3s; width: 100%; font-weight: bold; box-shadow: 0 4px 6px rgba(59,130,246,0.2); }
            .submit-btn:hover { background-color: #2563eb; transform: translateY(-1px); }
            .submit-btn:disabled { background-color: #94a3b8; cursor: not-allowed; transform: none; box-shadow: none; }

            #status { margin-top: 15px; color: #1e293b; font-size: 15px; font-weight: bold; min-height: 22px; }
            .example-text { font-size: 12px; color: #94a3b8; margin-top: 5px; line-height: 1.4; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 智能记账助手</h1>
            <div class="subtitle">让每一笔花销都清清楚楚</div>

            <div class="grid-container">
                <!-- 左侧：账单上传 -->
                <div class="card">
                    <h3>📁 1. 上传账单</h3>
                    <p style="font-size:12px; color:#64748b; margin:0 0 10px 0;">支持多选，可同时上传微信和支付宝</p>
                    <input type="file" id="billFiles" class="file-input" accept=".csv, .xlsx, .xls" multiple>
                </div>

                <!-- 右侧：分类规则 -->
                <div class="card">
                    <h3>🏷️ 2. 自定义分类 (选填)</h3>
                    <p style="font-size:12px; color:#64748b; margin:0 0 10px 0;">输入规则，让冷冰冰的“商户消费”变成你的专属分类：</p>
                    <textarea id="ruleText" class="rule-input" placeholder="例如：&#10;萧峰=人情代付&#10;星巴克=咖啡&#10;滴滴=交通"></textarea>
                    <div class="example-text">左边写关键词，右边写分类，用等号或冒号连接。</div>

                    <div style="margin-top: 15px; display: flex; align-items: center; justify-content: space-between;">
                        <span style="font-size: 12px; color: #64748b;">或者直接传 Excel：</span>
                        <a href="/download_template" class="btn-link">⬇️ 下载模板</a>
                    </div>
                    <input type="file" id="ruleFile" class="file-input" accept=".xlsx, .xls">
                </div>
            </div>

            <button class="submit-btn" id="submitBtn" onclick="uploadFiles()">一键生成记账表</button>
            <div id="status"></div>
        </div>

        <script>
            async function uploadFiles() {
                const billFiles = document.getElementById('billFiles').files;
                const ruleFile = document.getElementById('ruleFile').files[0];
                const ruleText = document.getElementById('ruleText').value;
                const statusDiv = document.getElementById('status');
                const submitBtn = document.getElementById('submitBtn');

                if (billFiles.length === 0) {
                    statusDiv.innerText = "⚠️ 请先在左边选择账单文件！";
                    statusDiv.style.color = "#dc2626";
                    return;
                }

                const formData = new FormData();
                for (let i = 0; i < billFiles.length; i++) {
                    formData.append("bill_files", billFiles[i]);
                }
                if (ruleFile) {
                    formData.append("rule_file", ruleFile);
                }
                if (ruleText.trim()) {
                    formData.append("rule_text", ruleText);
                }

                statusDiv.innerText = `正在处理中，请稍候...`;
                statusDiv.style.color = "#1e293b";
                submitBtn.disabled = true; submitBtn.innerText = "处理中...";

                try {
                    const response = await fetch('/generate', { method: 'POST', body: formData });
                    if (!response.ok) {
                        const errText = await response.text();
                        throw new Error(errText);
                    }
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url; a.download = "我的消费记账表.xlsx";
                    document.body.appendChild(a); a.click(); a.remove();
                    statusDiv.innerText = "✅ 完美！已开始下载。";
                    statusDiv.style.color = "#16a34a";
                } catch (error) {
                    statusDiv.innerText = "❌ " + error.message;
                    statusDiv.style.color = "#dc2626";
                } finally {
                    submitBtn.disabled = false; submitBtn.innerText = "一键生成记账表";
                }
            }
        </script>
    </body>
    </html>
    """


@app.post("/generate")
async def generate_report(
        background_tasks: BackgroundTasks,
        bill_files: list[UploadFile] = File(...),
        rule_file: UploadFile = File(None),
        rule_text: str = Form(None)
):
    temp_dir = f"temp_{uuid.uuid4().hex}"
    os.makedirs(temp_dir, exist_ok=True)

    try:
        user_mapping = {}

        if rule_text:
            lines = rule_text.strip().split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue
                parts = line.replace(':', '=').replace('：', '=').replace(' ', '=').split('=')
                if len(parts) >= 2:
                    key = parts[0].strip()
                    val = parts[1].strip()
                    if key and val:
                        user_mapping[key] = val

        if rule_file and rule_file.filename:
            rule_path = os.path.join(temp_dir, "rule.xlsx")
            with open(rule_path, "wb") as buffer:
                shutil.copyfileobj(rule_file.file, buffer)
            rule_df = pd.read_excel(rule_path)
            for _, row in rule_df.iterrows():
                key = str(row.iloc[0]).strip()
                val = str(row.iloc[1]).strip()
                if key and val and key != "nan" and val != "nan":
                    user_mapping[key] = val

        final_mapping = {**DEFAULT_MAPPING, **user_mapping}

        frames = []
        for file in bill_files:
            filepath = os.path.join(temp_dir, file.filename)
            with open(filepath, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            frames.append(process_uploaded_file(filepath, file.filename, final_mapping))

        all_df = pd.concat(frames, ignore_index=True).dropna(subset=["日期", "金额"])
        final_df = cancel_out_transactions(all_df)

        if final_df.empty: raise ValueError("没有找到有效的支出记录。")

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
            final_df[["日期", "月份", "消费类型", "金额", "来源"]].to_excel(writer, sheet_name="真实支出明细",
                                                                            index=False)
            auto_adjust_column_width(writer)
            format_excel(writer)

        background_tasks.add_task(shutil.rmtree, temp_dir, True)
        return FileResponse(output_path, filename="我的消费记账表.xlsx",
                            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        shutil.rmtree(temp_dir, True)
        return HTMLResponse(content=f"处理出错：{str(e)}", status_code=500)


if __name__ == "__main__":
    import uvicorn

    print("🚀 网页版记账助手已启动！")
    print("👉 请在浏览器中打开：http://127.0.0.1:8000")
    uvicorn.run(app, host="0.0.0.0", port=8000)
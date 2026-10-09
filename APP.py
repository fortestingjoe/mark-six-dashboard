import io
import random
from collections import Counter

import dash
from dash import Dash, Input, Output, State, dcc, html, dash_table
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px

# 1. 歷史 20 期六合彩資料
HISTORY_DRAWS = [
    ("2026/10/08", [18, 19, 24, 31, 40, 44], 22),
    ("2026/10/06", [9, 29, 31, 34, 36, 41], 28),
    ("2026/10/03", [8, 23, 29, 31, 34, 43], 48),
    ("2026/09/26", [4, 28, 31, 44, 47, 48], 19),
    ("2026/09/22", [2, 34, 35, 43, 45, 46], 41),
    ("2026/09/19", [5, 8, 11, 14, 19, 31], 15),
    ("2026/09/15", [7, 12, 18, 21, 33, 49], 11),
    ("2026/09/10", [15, 22, 23, 31, 34, 35], 45),
    ("2026/09/08", [9, 23, 28, 29, 35, 41], 38),
    ("2026/09/05", [9, 18, 26, 30, 33, 45], 42),
    ("2026/08/27", [12, 17, 25, 31, 44, 45], 35),
    ("2026/08/25", [1, 18, 19, 25, 34, 38], 29),
    ("2026/08/22", [7, 9, 12, 25, 34, 40], 29),
    ("2026/08/20", [7, 9, 10, 15, 24, 46], 34),
    ("2026/08/18", [7, 8, 9, 14, 39, 41], 49),
    ("2026/08/15", [4, 16, 25, 27, 28, 33], 14),
    ("2026/08/11", [2, 15, 18, 23, 39, 41], 47),
    ("2026/08/06", [13, 24, 29, 30, 42, 48], 3),
    ("2026/08/04", [4, 11, 13, 16, 31, 33], 38),
    ("2026/08/01", [1, 7, 16, 22, 32, 37], 48),
]

# 2. 數據整理與統計
table_rows = []
all_main_numbers = []

for date_str, nums, extra in HISTORY_DRAWS:
  all_main_numbers.extend(nums)
  table_rows.append({
      "date": date_str,
      "n1": nums[0],
      "n2": nums[1],
      "n3": nums[2],
      "n4": nums[3],
      "n5": nums[4],
      "n6": nums[5],
      "s_num": extra,
  })

df_draws = pd.DataFrame(table_rows)

# 統計出現次數
counts = Counter(all_main_numbers)
df_counts = pd.DataFrame(
    [
        {"號碼": f"{num:02d}", "出現次數": counts.get(num, 0)}
        for num in range(1, 50)
    ]
)
df_top10 = df_counts.sort_values(
    by="出現次數", ascending=False
).head(10)

# 冷門號碼 (0次)
cold_nums = [
    f"{num:02d}" for num in range(1, 50) if counts.get(num, 0) == 0
]

# 3. 建立 Dash App
app = Dash(__name__)
app.title = "六合彩 Dashboard"
server = app.server  # 👈 新增這一行（讓 Render/Gunicorn 讀取）

# ...中間程式碼保持不變...

if __name__ == "__main__":
  app.run(debug=True)

# 4. 畫圖 (Top 10 熱門號碼柱狀圖)
fig_chart = px.bar(
    df_top10,
    x="號碼",
    y="出現次數",
    title="🔥 近 20 期熱門號碼 Top 10 (主要號碼)",
    text="出現次數",
    color="出現次數",
    color_continuous_scale="Reds",
)
fig_chart.update_layout(
    xaxis_type="category",
    template="plotly_white",
    margin=dict(l=20, r=20, t=50, b=20),
)

# 5. Dashboard 介面佈局
app.layout = html.Div(
    style={
        "backgroundColor": "#f4f6f9",
        "fontFamily": "Microsoft JhengHei, sans-serif",
        "padding": "20px",
    },
    children=[
        # 標題列
        html.Div(
            style={
                "backgroundColor": "#1F497D",
                "color": "white",
                "padding": "15px 20px",
                "borderRadius": "8px",
                "marginBottom": "20px",
            },
            children=[
                html.H2(
                    "🎯 香港六合彩 (Mark Six) 數據分析 Dashboard",
                    style={"margin": 0},
                )
            ],
        ),
        # 區塊 1：1-49 隨機出號器
        html.Div(
            style={
                "backgroundColor": "white",
                "padding": "20px",
                "borderRadius": "8px",
                "boxShadow": "0 2px 4px rgba(0,0,0,0.1)",
                "marginBottom": "20px",
            },
            children=[
                html.H4("🎲 1 - 49 隨機攪珠選號", style={"marginTop": 0}),
                html.Button(
                    "隨機產生 6 個號碼",
                    id="btn-generate",
                    n_clicks=0,
                    style={
                        "backgroundColor": "#28a745",
                        "color": "white",
                        "border": "none",
                        "padding": "10px 20px",
                        "fontSize": "16px",
                        "borderRadius": "5px",
                        "cursor": "pointer",
                    },
                ),
                html.Div(
                    id="gen-result",
                    style={
                        "marginTop": "15px",
                        "fontSize": "20px",
                        "fontWeight": "bold",
                        "color": "#28a745",
                    },
                ),
            ],
        ),
        # 區塊 2：冷熱門數據卡片 & 圖表
        html.Div(
            style={
                "display": "flex",
                "gap": "20px",
                "marginBottom": "20px",
                "flexWrap": "wrap",
            },
            children=[
                # 圖表
                html.Div(
                    style={
                        "flex": "2",
                        "backgroundColor": "white",
                        "padding": "15px",
                        "borderRadius": "8px",
                        "boxShadow": "0 2px 4px rgba(0,0,0,0.1)",
                        "minWidth": "350px",
                    },
                    children=[dcc.Graph(figure=fig_chart)],
                ),
                # 冷門號碼卡片
                html.Div(
                    style={
                        "flex": "1",
                        "backgroundColor": "white",
                        "padding": "20px",
                        "borderRadius": "8px",
                        "boxShadow": "0 2px 4px rgba(0,0,0,0.1)",
                        "minWidth": "250px",
                    },
                    children=[
                        html.H4(
                            "❄️ 未曾開出號碼 (冷門)",
                            style={"color": "#0275d8", "marginTop": 0},
                        ),
                        html.P("近 20 期主要攪珠未開出："),
                        html.H3(
                            (
                                ", ".join(cold_nums)
                                if cold_nums
                                else "所有號碼皆有開出"
                            ),
                            style={"color": "#d9534f"},
                        ),
                    ],
                ),
            ],
        ),
        # 區塊 3：20 期詳細歷史紀錄表格
        html.Div(
            style={
                "backgroundColor": "white",
                "padding": "20px",
                "borderRadius": "8px",
                "boxShadow": "0 2px 4px rgba(0,0,0,0.1)",
                "marginBottom": "20px",
            },
            children=[
                html.Div(
                    style={
                        "display": "flex",
                        "justifyContent": "space-between",
                        "alignItems": "center",
                        "marginBottom": "15px",
                    },
                    children=[
                        html.H4(
                            "📋 近 20 期詳細開獎紀錄", style={"margin": 0}
                        ),
                        html.Button(
                            "📊 下載 Excel 分析報表",
                            id="btn-excel",
                            n_clicks=0,
                            style={
                                "backgroundColor": "#17a2b8",
                                "color": "white",
                                "border": "none",
                                "padding": "10px 18px",
                                "fontSize": "14px",
                                "borderRadius": "5px",
                                "cursor": "pointer",
                            },
                        ),
                        dcc.Download(id="download-excel"),
                    ],
                ),
                dash_table.DataTable(
                    data=df_draws.to_dict("records"),
                    columns=[
                        {"name": "期數/日期", "id": "date"},
                        {"name": "號碼 1", "id": "n1"},
                        {"name": "號碼 2", "id": "n2"},
                        {"name": "號碼 3", "id": "n3"},
                        {"name": "號碼 4", "id": "n4"},
                        {"name": "號碼 5", "id": "n5"},
                        {"name": "號碼 6", "id": "n6"},
                        {"name": "特別號", "id": "s_num"},
                    ],
                    style_header={
                        "backgroundColor": "#1F497D",
                        "color": "white",
                        "fontWeight": "bold",
                        "textAlign": "center",
                    },
                    style_cell={"textAlign": "center", "padding": "8px"},
                    style_data_conditional=[
                        {
                            "if": {"row_index": "odd"},
                            "backgroundColor": "#f9fbfd",
                        },
                        {
                            "if": {"column_id": "s_num"},
                            "color": "#C00000",
                            "fontWeight": "bold",
                        },
                    ],
                ),
            ],
        ),
    ],
)


# 6. Callbacks (互動邏輯)
# 邏輯 A: 隨機選號
@app.callback(
    Output("gen-result", "children"), Input("btn-generate", "n_clicks")
)
def generate_random_numbers(n_clicks):
  if n_clicks == 0:
    return "點擊按鈕開始選號..."
  nums = sorted(random.sample(range(1, 50), 6))
  return f"抽中號碼：[ {'  .  '.join([f'{n:02d}' for n in nums])} ]"


# 邏輯 B: 匯出 Excel 下載
@app.callback(
    Output("download-excel", "data"),
    Input("btn-excel", "n_clicks"),
    prevent_initial_call=True,
)
def export_excel(n_clicks):
  wb = openpyxl.Workbook()

  ws_summary = wb.active
  ws_summary.title = "統計摘要"
  ws_summary.views.sheetView[0].showGridLines = True

  ws_detail = wb.create_sheet(title="近20期詳細紀錄")
  ws_detail.views.sheetView[0].showGridLines = True

  font_title = Font(
      name="Microsoft JhengHei", size=15, bold=True, color="1F497D"
  )
  font_header = Font(
      name="Microsoft JhengHei", size=11, bold=True, color="FFFFFF"
  )
  font_body = Font(name="Microsoft JhengHei", size=11)
  font_bold = Font(name="Microsoft JhengHei", size=11, bold=True)
  font_special = Font(
      name="Microsoft JhengHei", size=11, bold=True, color="C00000"
  )

  fill_header = PatternFill(
      start_color="1F497D", end_color="1F497D", fill_type="solid"
  )
  fill_zebra = PatternFill(
      start_color="F2F5F9", end_color="F2F5F9", fill_type="solid"
  )
  fill_top3 = PatternFill(
      start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
  )

  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )
  align_center = Alignment(horizontal="center", vertical="center")

  # 寫入 Detail 頁籤
  ws_detail.cell(
      row=1, column=1, value="香港六合彩 - 近20期詳細開獎紀錄"
  ).font = font_title
  headers_detail = [
      "期數/日期",
      "號碼 1",
      "號碼 2",
      "號碼 3",
      "號碼 4",
      "號碼 5",
      "號碼 6",
      "特別號",
  ]
  for c_idx, text in enumerate(headers_detail, 1):
    cell = ws_detail.cell(row=3, column=c_idx, value=text)
    cell.font, cell.fill, cell.alignment, cell.border = (
        font_header,
        fill_header,
        align_center,
        thin_border,
    )

  for r_idx, (date_str, nums, extra) in enumerate(HISTORY_DRAWS, 4):
    row_vals = [date_str] + nums + [extra]
    for c_idx, val in enumerate(row_vals, 1):
      cell = ws_detail.cell(row=r_idx, column=c_idx, value=val)
      cell.font = font_special if c_idx == 8 else font_body
      cell.alignment, cell.border = align_center, thin_border
      if r_idx % 2 == 0:
        cell.fill = fill_zebra

  # 寫入 Summary 頁籤
  ws_summary.cell(
      row=1, column=1, value="香港六合彩 - 近20期出現頻率統計分析"
  ).font = font_title
  ws_summary.cell(
      row=3, column=1, value="【1-49號碼次數統計表】"
  ).font = font_bold
  headers_freq = ["號碼", "出現次數 (主號)", "出現率"]
  for c_idx, text in enumerate(headers_freq, 1):
    cell = ws_summary.cell(row=4, column=c_idx, value=text)
    cell.font, cell.fill, cell.alignment, cell.border = (
        font_header,
        fill_header,
        align_center,
        thin_border,
    )

  for num in range(1, 50):
    r_idx = num + 4
    c_num = ws_summary.cell(row=r_idx, column=1, value=num)
    c_cnt = ws_summary.cell(
        row=r_idx,
        column=2,
        value=f"=COUNTIF('近20期詳細紀錄'!$B$4:$G$23, A{r_idx})",
    )
    c_pct = ws_summary.cell(row=r_idx, column=3, value=f"=B{r_idx}/20")
    c_pct.number_format = "0.0%"
    for c in [c_num, c_cnt, c_pct]:
      c.font, c.alignment, c.border = font_body, align_center, thin_border
      if num % 2 == 0:
        c.fill = fill_zebra

  ws_summary.cell(row=3, column=5, value="🔥【最熱門號碼 Top 10】").font = (
      font_bold
  )
  headers_top = ["名次", "熱門號碼", "出現次數"]
  for c_idx, text in enumerate(headers_top, 5):
    cell = ws_summary.cell(row=4, column=c_idx, value=text)
    cell.font, cell.fill, cell.alignment, cell.border = (
        font_header,
        fill_header,
        align_center,
        thin_border,
    )

  top_10 = Counter(all_main_numbers).most_common(10)
  for idx, (num, cnt) in enumerate(top_10, 1):
    r_idx = idx + 4
    cells = [
        ws_summary.cell(row=r_idx, column=5, value=f"第 {idx} 名"),
        ws_summary.cell(row=r_idx, column=6, value=num),
        ws_summary.cell(row=r_idx, column=7, value=cnt),
    ]
    for c in cells:
      c.font = font_bold if idx <= 3 else font_body
      c.alignment, c.border = align_center, thin_border
      if idx <= 3:
        c.fill = fill_top3
      elif idx % 2 == 0:
        c.fill = fill_zebra

  for ws in [ws_summary, ws_detail]:
    for col in ws.columns:
      max_len = max(len(str(cell.value or "")) for cell in col)
      col_letter = get_column_letter(col[0].column)
      ws.column_dimensions[col_letter].width = max(max_len + 5, 12)

  # 將檔案寫入記憶體流供瀏覽器下載
  output = io.BytesIO()
  wb.save(output)
  output.seek(0)

  return dcc.send_bytes(output.getvalue(), filename="MarkSix_Dashboard_Report.xlsx")


# 7. 啟動伺服器
if __name__ == "__main__":
  app.run(debug=True)
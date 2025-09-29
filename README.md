# Employee Performance Analysis System

Hệ thống phân tích hiệu suất nhân viên toàn diện với AI và phân tích 360 độ.

## Tính năng chính

- **Phân tích theo thời gian**: Đánh giá hiệu suất nhân viên theo kỳ (tháng/quý/năm)
- **Employee 360**: Phân tích toàn diện về kỹ năng, hợp tác, lãnh đạo, và ảnh hưởng
- **Phân tích dự án**: Đánh giá chi tiết timeline, chất lượng, rủi ro của dự án
- **Phân tích AI**: Sử dụng Google Gemini/Gemma cho insights thông minh
- **Phân tích cơ bản**: Thống kê và metrics không cần AI

## Cài đặt

```bash
pip install streamlit pandas numpy plotly requests beautifulsoup4 python-docx
pip install google-genai  # Optional: for AI analysis
```

## Cấu hình

Cần thiết lập các biến môi trường hoặc cập nhật trong code:

```python
WEWORK_ACCESS_TOKEN = "your_wework_token"
ACCOUNT_ACCESS_TOKEN = "your_account_token"
GEMINI_API_KEY = "your_gemini_api_key"  # Optional
```

## Sử dụng

```bash
streamlit run app.py
```

## Chức năng

### 1. Phân tích theo Thời gian
- Đánh giá hiệu suất trong kỳ
- So sánh tiến độ hoàn thành
- Phân tích chất lượng dữ liệu

### 2. Employee 360
- Networking & collaboration score
- Skill matrix & expertise level
- Leadership & influence analysis
- Growth trajectory & innovation
- Peer comparison

### 3. Phân tích Dự án
- Overview dashboard
- Timeline & deadline tracking
- Quality metrics
- Risk assessment
- Collaboration analysis

## Export

Hỗ trợ xuất báo cáo:
- Word (.docx)
- Excel (.xlsx)
- CSV
- JSON
- PDF (cần reportlab)

## Công nghệ

- **Frontend**: Streamlit
- **Visualization**: Plotly, Matplotlib, Seaborn
- **AI**: Google Gemini 2.0 Flash / Gemma 3
- **API**: WeWork, Account Base.vn
- **Export**: python-docx, pandas, openpyxl

## Lưu ý

- AI analysis yêu cầu Google GenAI API key
- PDF export cần cài thêm: `pip install reportlab`
- Cần quyền truy cập WeWork và Account APIs

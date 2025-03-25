# Google Sheets API

API này cung cấp các phương thức để thao tác với Google Sheets thông qua các phương thức HTTP (`GET`, `POST`, `PUT`, `DELETE`).

## 🚀 Hướng dẫn cài đặt

1. **Triển khai Google Apps Script:**

   - Mở Google Sheets
   - Vào `Extensions` > `Apps Script`
   - Sao chép `api.gs` và dán code vào
   - Triển khai script ở chế độ Web App
   - Lưu lại URL của script để sử dụng

2. **Sử dụng API**
   - Tất cả các yêu cầu được gửi đến URL của Google Apps Script

## 📌 API Endpoints

### 1️⃣ Lấy dữ liệu từ Sheet (`GET`)

**Endpoint:**

```plaintext
GET <YOUR_SCRIPT_URL>?gid=<SHEET_ID>
```

**Request Parameters:**

- `gid` (bắt buộc): ID của sheet cần lấy dữ liệu.

**Response Example:**

```json
{
  "sheet_name": "Sheet1",
  "data": [
    ["Name", "Email"],
    ["Alice", "alice@example.com"]
  ],
  "backgrounds": [
    ["#ffffff", "#ffffff"],
    ["#00ffff", "#ffffff"]
  ]
}
```

---

### 2️⃣ Thêm dữ liệu vào Sheet (`POST`)

**Endpoint:**

```plaintext
POST <YOUR_SCRIPT_URL>
```

**Request Body:**

```json
{
  "method": "POST",
  "sheet_id": "2063363012",
  "data": ["John Doe", "john@example.com"]
}
```

**Response:**

```json
{ "success": "ok" }
```

---

### 3️⃣ Xóa dòng trong Sheet (`DELETE`)

**Endpoint:**

```plaintext
POST <YOUR_SCRIPT_URL>
```

**Request Body:**

```json
{
  "method": "DELETE",
  "sheet_id": "2063363012",
  "row": 3
}
```

**Response:**

```json
{ "success": "ok" }
```

---

### 4️⃣ Cập nhật dữ liệu và màu nền (`PUT`)

**Endpoint:**

```plaintext
POST <YOUR_SCRIPT_URL>
```

**Request Body:**

```json
{
  "method": "PUT",
  "sheet_id": "2063363012",
  "row": 2,
  "data": ["Updated Name", "updated@example.com"],
  "color": "#ffff00"
}
```

**Response:**

```json
{ "success": "ok" }
```

## 🔹 Lưu ý

- `sheet_id`: ID của sheet (có thể lấy từ Google Apps Script)
- `row`: Dòng trong sheet (bắt đầu từ 1, không tính tiêu đề)
- `data`: Danh sách giá trị cần thêm hoặc cập nhật
- `color`: Mã màu HEX để tô nền

## 📌 Liên hệ

Nếu gặp vấn đề hoặc cần hỗ trợ, hãy liên hệ [your-email@example.com]

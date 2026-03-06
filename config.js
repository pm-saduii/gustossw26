// config.js
// 1. นำ URL ที่ได้จากตอน Deploy Google Apps Script (Web App) มาใส่ที่นี่
const API_URL = 'https://script.google.com/macros/s/AKfycbwQbc31pY86tM7EdD41dBsXRQ-OmL6M5m0TJKs_ZJiix2bLEYkHOiQ8eIGazR7ay0_l/exec'; 

// 2. ฟังก์ชันกลางสำหรับการเรียก API ที่รองรับการส่ง Token อัตโนมัติ
async function apiCall(params) {
    const token = localStorage.getItem('vms_token');
    const urlParams = new URLSearchParams(params);
    
    // ส่ง Token ไปกับทุก Request เพื่อตรวจสอบสิทธิ์ในฝั่ง Server
    if (token) urlParams.append('token', token);

    try {
        const response = await fetch(`${API_URL}?${urlParams.toString()}`);
        if (!response.ok) throw new Error('Network response was not ok');
        
        const data = await response.json();
        
        // ถ้า Token หมดอายุ หรือไม่ถูกต้อง ให้เตะกลับไปหน้า Login
        if (data.error === 'Unauthorized') {
            localStorage.clear();
            window.location.href = 'index.html';
            return null;
        }
        return data;
    } catch (error) {
        console.error('API Error:', error);
        // แสดง Error แบบนุ่มนวล
        if (typeof Swal !== 'undefined') {
            Swal.fire({ icon: 'error', title: 'การเชื่อมต่อขัดข้อง', text: 'ไม่สามารถติดต่อเซิร์ฟเวอร์ได้ในขณะนี้' });
        }
        return { success: false, message: 'การเชื่อมต่อเซิร์ฟเวอร์ขัดข้อง' };
    }
}
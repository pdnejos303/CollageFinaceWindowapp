import markdown
import os

with open('DOCUMENTATION.md', 'r', encoding='utf-8') as f:
    md_content = f.read()

# Convert markdown to html
html_body = markdown.markdown(md_content, extensions=['tables', 'fenced_code', 'toc'])

html_template = """<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>เอกสารประกอบโครงการ & คู่มือการใช้งานระบบ SPVC Smart Receipt</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #0f172a;
            --primary-light: #1e293b;
            --accent: #2563eb;
            --accent-hover: #1d4ed8;
            --success: #059669;
            --danger: #dc2626;
            --text-main: #1e293b;
            --text-muted: #64748b;
            --bg-body: #f8fafc;
            --bg-card: #ffffff;
            --border: #cbd5e1;
            --border-light: #e2e8f0;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: 'Sarabun', -apple-system, BlinkMacSystemFont, 'Segoe UI', Tahoma, sans-serif;
            background-color: var(--bg-body);
            color: var(--text-main);
            line-height: 1.7;
            font-size: 15px;
        }

        header.top-header {
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
            color: white;
            padding: 30px 40px;
            border-bottom: 4px solid var(--accent);
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }

        header.top-header h1 {
            font-size: 26px;
            font-weight: 700;
            margin-bottom: 8px;
            color: #ffffff;
        }

        header.top-header p {
            font-size: 15px;
            color: #94a3b8;
        }

        .action-bar {
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            margin-top: 15px;
        }

        .btn {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 9px 18px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            text-decoration: none;
            transition: all 0.2s;
            border: none;
            font-family: inherit;
        }

        .btn-primary {
            background: var(--accent);
            color: white;
        }
        .btn-primary:hover {
            background: var(--accent-hover);
        }

        .btn-secondary {
            background: #334155;
            color: white;
        }
        .btn-secondary:hover {
            background: #475569;
        }

        .container {
            max-width: 1350px;
            margin: 30px auto;
            padding: 0 25px;
            display: grid;
            grid-template-columns: 280px 1fr;
            gap: 30px;
        }

        @media (max-width: 1024px) {
            .container {
                grid-template-columns: 1fr;
            }
            .sidebar {
                display: none;
            }
        }

        .sidebar {
            position: sticky;
            top: 25px;
            height: calc(100vh - 50px);
            overflow-y: auto;
            background: var(--bg-card);
            padding: 20px;
            border: 1px solid var(--border-light);
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }

        .sidebar h3 {
            font-size: 16px;
            margin-bottom: 15px;
            color: var(--primary);
            border-bottom: 2px solid var(--accent);
            padding-bottom: 8px;
        }

        .sidebar ul {
            list-style: none;
        }

        .sidebar li {
            margin-bottom: 6px;
        }

        .sidebar a {
            color: #475569;
            text-decoration: none;
            font-size: 13.5px;
            display: block;
            padding: 6px 10px;
            border-radius: 4px;
            transition: all 0.15s;
        }

        .sidebar a:hover {
            background: #f1f5f9;
            color: var(--accent);
            font-weight: 600;
            padding-left: 14px;
        }

        .main-content {
            background: var(--bg-card);
            padding: 40px;
            border: 1px solid var(--border-light);
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
            overflow-x: auto;
        }

        .main-content h1, .main-content h2, .main-content h3, .main-content h4 {
            color: var(--primary);
            margin-top: 35px;
            margin-bottom: 15px;
            font-weight: 700;
        }

        .main-content h1 {
            font-size: 26px;
            border-bottom: 2px solid var(--border-light);
            padding-bottom: 10px;
        }

        .main-content h2 {
            font-size: 20px;
            border-left: 5px solid var(--accent);
            padding-left: 14px;
            background: #f8fafc;
            padding-top: 8px;
            padding-bottom: 8px;
        }

        .main-content h3 {
            font-size: 17px;
            color: #1e293b;
            margin-top: 25px;
        }

        .main-content h4 {
            font-size: 15px;
            color: #1e293b;
            margin-top: 22px;
            background: #f1f5f9;
            display: inline-block;
            padding: 5px 14px;
            border-radius: 4px;
            border-left: 3px solid #2563eb;
        }

        .main-content p {
            margin-bottom: 16px;
            color: #334155;
        }

        .main-content ul, .main-content ol {
            margin-bottom: 20px;
            padding-left: 28px;
            color: #334155;
        }

        .main-content li {
            margin-bottom: 6px;
        }

        /* TABLE STYLING - MODERN CORP */
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0 30px 0;
            font-size: 14px;
            background: white;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.03);
        }

        thead {
            background: #0f172a;
        }

        th {
            background: #0f172a;
            color: #ffffff;
            font-weight: 600;
            text-align: left;
            padding: 12px 14px;
            border: 1px solid #334155;
            letter-spacing: 0.2px;
        }

        td {
            padding: 10px 14px;
            border: 1px solid #e2e8f0;
            color: #334155;
            vertical-align: top;
        }

        tbody tr:nth-child(even) {
            background-color: #f8fafc;
        }

        tbody tr:hover {
            background-color: #f1f5f9;
        }

        code {
            font-family: 'JetBrains Mono', Consolas, monospace;
            background: #f1f5f9;
            color: #0f172a;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 13px;
            border: 1px solid #e2e8f0;
        }

        pre {
            background: #0f172a;
            color: #f8fafc;
            padding: 18px 20px;
            border-radius: 8px;
            overflow-x: auto;
            margin: 20px 0;
            font-family: 'JetBrains Mono', Consolas, monospace;
            font-size: 13.5px;
            line-height: 1.6;
        }

        pre code {
            background: transparent;
            color: inherit;
            padding: 0;
            border: none;
        }

        blockquote {
            border-left: 4px solid var(--accent);
            padding: 12px 20px;
            background: #eff6ff;
            color: #1e3a8a;
            margin: 20px 0;
            border-radius: 0 6px 6px 0;
        }

        hr {
            border: none;
            border-top: 1px solid var(--border-light);
            margin: 35px 0;
        }

        @media print {
            .sidebar, .action-bar, header.top-header {
                display: none;
            }
            .container {
                grid-template-columns: 1fr;
                max-width: 100%;
                margin: 0;
                padding: 0;
            }
            .main-content {
                border: none;
                box-shadow: none;
                padding: 0;
            }
            table {
                page-break-inside: avoid;
            }
        }
    </style>
</head>
<body>

<header class="top-header">
    <h1>📘 SPVC Smart Receipt Desktop System</h1>
    <p>วิทยาลัยอาชีวศึกษาสุพรรณบุรี • เอกสารประกอบโครงการ สถาปัตยกรรมระบบ โครงสร้างฐานข้อมูล และคู่มือการใช้งานฉบับสมบูรณ์</p>
    <div class="action-bar">
        <button class="btn btn-primary" onclick="window.print()">🖨️ สั่งพิมพ์ / บันทึกเป็น PDF</button>
        <a class="btn btn-secondary" href="#4-ระบบฐานข้อมูล-โครงสร้างตาราง-และตำแหน่งที่ตั้ง-database-architecture">🗄️ ดูตารางฐานข้อมูลทั้ง 11 ตาราง</a>
        <a class="btn btn-secondary" href="#6-คู่มือการใช้งานโปรแกรมฉบับสมบูรณ์-user-manual">📖 ไปที่คู่มือการใช้งาน</a>
    </div>
</header>

<div class="container">
    <nav class="sidebar">
        <h3>📑 สารบัญเอกสาร</h3>
        <ul>
            <li><a href="#1-ภาพรวมของโปรเจค-project-overview">1. ภาพรวมของโปรเจค</a></li>
            <li><a href="#2-เทคโนโลยีและไลบรารีที่ใช้-tech-stack">2. เทคโนโลยีและ Tech Stack</a></li>
            <li><a href="#3-โครงสร้างโปรเจค-project-structure">3. โครงสร้างโปรเจค</a></li>
            <li><a href="#4-ระบบฐานข้อมูล-โครงสร้างตาราง-และตำแหน่งที่ตั้ง-database-architecture">4. สถาปัตยกรรมฐานข้อมูล</a></li>
            <li><a href="#43-ตารางภาพรวมฐานข้อมูล-database-schema-overview---11-ตาราง" style="padding-left: 20px;">↳ ภาพรวม 11 ตาราง</a></li>
            <li><a href="#44-โครงสร้างคอลัมน์ของแต่ละตารางอย่างละเอียด-detailed-table-schemas" style="padding-left: 20px;">↳ สคีมาละเอียดทุกตาราง</a></li>
            <li><a href="#45-ดัชนีเพื่อเพิ่มประสิทธิภาพการสืบค้น-database-indexes---8-indexes" style="padding-left: 20px;">↳ ตาราง Index 8 ตัว</a></li>
            <li><a href="#5-ระบบความปลอดภัยและการตรวจสอบย้อนหลัง-security--compliance">5. ระบบความปลอดภัย & สิทธิ์</a></li>
            <li><a href="#6-คู่มือการใช้งานโปรแกรมฉบับสมบูรณ์-user-manual">6. คู่มือการใช้งานโปรแกรม</a></li>
            <li><a href="#7-การแก้ไขปัญหาเบื้องต้น-troubleshooting--faq">7. การแก้ไขปัญหา (FAQ)</a></li>
            <li><a href="#8-มาตรฐานการออกแบบส่วนต่อประสานผู้ใช้-clean-ui--single-enclosure-rule">8. มาตรฐานการออกแบบ UI</a></li>
        </ul>
    </nav>

    <main class="main-content">
""" + html_body + """
    </main>
</div>

</body>
</html>
"""

with open('DOCUMENTATION.html', 'w', encoding='utf-8') as f:
    f.write(html_template)

print('Successfully generated DOCUMENTATION.html!')

# -*- coding: utf-8 -*-
"""
内网文档处理工作台 — 后端服务

运行方式:
    源码运行   python app.py
    打包运行   内网文档处理工作台.exe

访问地址: http://127.0.0.1:5000
依赖: flask, python-docx, openpyxl, python-pptx, PyPDF2
"""

import os
import re
import sys
import json
import shutil
import socket
import zipfile
import logging
import threading
import webbrowser
from datetime import datetime, date
from copy import deepcopy

from flask import Flask, request, jsonify, send_file, render_template, abort
from werkzeug.serving import run_simple

from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from pptx import Presentation
from pptx.util import Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from PyPDF2 import PdfMerger


APP_NAME = '内网文档处理工作台'
APP_VERSION = 'v1.0'


# ============================================================
# 1. 应用初始化
# ============================================================
def _is_frozen():
    """是否运行在 PyInstaller 打包出的可执行文件中"""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')


def _resource_dir():
    """只读资源目录（templates 等）"""
    if _is_frozen():
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def _data_dir():
    """可写数据目录（config.json / uploads / downloads）"""
    if _is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


RESOURCE_DIR = _resource_dir()
DATA_DIR = _data_dir()

app = Flask(__name__, template_folder=os.path.join(RESOURCE_DIR, 'templates'))
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024
app.config['JSON_AS_ASCII'] = False
app.config['TEMPLATES_AUTO_RELOAD'] = True

UPLOAD_FOLDER = os.path.join(DATA_DIR, 'uploads')
DOWNLOAD_FOLDER = os.path.join(DATA_DIR, 'downloads')
CONFIG_FILE = os.path.join(DATA_DIR, 'config.json')

for _d in (UPLOAD_FOLDER, DOWNLOAD_FOLDER):
    os.makedirs(_d, exist_ok=True)


# ============================================================
# 2. 功能元数据
# ============================================================
FEATURE_META = [
    {
        "key": "word_merge", "category": "Word 文档深度处理",
        "icon": "📄", "name": "多文档一键合并",
        "desc": "上传多份 Word 文档，按上传顺序合并为一份，每份之间自动插入分页符。",
        "accept": ".docx", "multiple": True
    },
    {
        "key": "word_redact", "category": "Word 文档深度处理",
        "icon": "🛡️", "name": "文档脱敏与排版",
        "desc": "正则脱敏手机号/身份证号/座机，清除多余空行，全文段落首行缩进2字符。",
        "accept": ".docx", "multiple": False
    },
    {
        "key": "word_extract", "category": "Word 文档深度处理",
        "icon": "🔍", "name": "关键信息提取库",
        "desc": "批量提取时间、地点、电话、身份证、邮箱、金额等关键信息，汇总为Excel。",
        "accept": ".docx", "multiple": True
    },
    {
        "key": "excel_split", "category": "Excel台账智能清洗",
        "icon": "📊", "name": "大表按分类拆分",
        "desc": "按指定列(如\"部门\")的唯一值，将大表拆分为多个Excel，打包ZIP下载。",
        "accept": ".xlsx", "multiple": False,
        "params": [{"key": "column", "label": "拆分列名", "default": "部门", "type": "text"}]
    },
    {
        "key": "excel_merge", "category": "Excel台账智能清洗",
        "icon": "🔗", "name": "多表横向对齐合并",
        "desc": "以首文件为主表，按关键列(如身份证号)横向拼接其他表的所有列。",
        "accept": ".xlsx", "multiple": True,
        "params": [{"key": "keycol", "label": "匹配列名", "default": "身份证号", "type": "text"}]
    },
    {
        "key": "excel_idcard", "category": "Excel台账智能清洗",
        "icon": "🆔", "name": "身份证信息智能解析",
        "desc": "自动识别身份证列，新增\"出生日期\"、\"性别\"、\"年龄\"三列并填充。",
        "accept": ".xlsx", "multiple": False,
        "params": [{"key": "idcol", "label": "身份证列名(留空自动识别)", "default": "", "type": "text"}]
    },
    {
        "key": "ppt_outline", "category": "PPT汇报结构化",
        "icon": "📑", "name": "汇报大纲一键提取",
        "desc": "提取每页PPT的标题与正文，按层级结构生成Word大纲文档。",
        "accept": ".pptx", "multiple": False
    },
    {
        "key": "ppt_watermark", "category": "PPT汇报结构化",
        "icon": "💧", "name": "PPT批量加防伪水印",
        "desc": "在每页PPT中央加入指定文字水印，支持自定义水印文案。",
        "accept": ".pptx", "multiple": True,
        "params": [{"key": "text", "label": "水印文字", "default": "内部资料 严禁外传", "type": "text"}]
    },
    {
        "key": "ppt_images", "category": "PPT汇报结构化",
        "icon": "🖼️", "name": "PPT图片提取打包",
        "desc": "一键打包PPT中嵌入的全部原始图片(媒体文件夹)下载。",
        "accept": ".pptx", "multiple": True
    },
    {
        "key": "pdf_merge", "category": "PDF与档案批处理",
        "icon": "📚", "name": "多PDF一键合并",
        "desc": "上传多份PDF，按上传顺序合并为一份新PDF。",
        "accept": ".pdf", "multiple": True
    },
    {
        "key": "file_rename", "category": "PDF与档案批处理",
        "icon": "🏷️", "name": "批量文件规范化重命名",
        "desc": "按\"前缀+原名+后缀\"规则批量重命名，支持任意文件类型。",
        "accept": "*", "multiple": True,
        "params": [
            {"key": "prefix", "label": "统一前缀", "default": "内网文件_", "type": "text"},
            {"key": "suffix", "label": "统一后缀", "default": "", "type": "text"}
        ]
    },
    {
        "key": "text_clean", "category": "PDF与档案批处理",
        "icon": "🧹", "name": "长文本极速清洗",
        "desc": "去除乱码/空行/控制字符，提取包含指定关键词的句子，纯本地零依赖。",
        "accept": ".txt", "multiple": False,
        "params": [
            {"key": "text", "label": "直接粘贴文本(也可上传txt)", "default": "", "type": "textarea"},
            {"key": "keyword", "label": "提取关键词(留空输出全部)", "default": "", "type": "text"}
        ]
    },
]


# ============================================================
# 3. 配置管理
# ============================================================
DEFAULT_CONFIG = {item["key"]: True for item in FEATURE_META}


def load_config():
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        return DEFAULT_CONFIG.copy()
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    for k, v in DEFAULT_CONFIG.items():
        cfg.setdefault(k, v)
    return cfg


def save_config(cfg):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def is_enabled(feature_key):
    return bool(load_config().get(feature_key, True))


# ============================================================
# 4. 通用工具
# ============================================================
def err(msg, status=400):
    return jsonify({"status": "error", "message": msg}), status


def gen_filename(prefix, ext):
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    rand = datetime.now().strftime('%f')[:3]
    return f"{prefix}_{ts}_{rand}.{ext}"


def save_uploads(files, sub=''):
    """保存上传文件到临时工作目录，返回 (工作目录, [文件路径])"""
    work_dir = os.path.join(UPLOAD_FOLDER, sub, datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    os.makedirs(work_dir, exist_ok=True)
    paths = []
    for idx, f in enumerate(files):
        if not f or not f.filename:
            continue
        original = os.path.basename(f.filename)
        original = re.sub(r'[\\/:*?"<>|]', '_', original)
        name = f"{idx:03d}_{original}"
        full = os.path.join(work_dir, name)
        f.save(full)
        paths.append(full)
    return work_dir, paths


def safe_clear(dirpath):
    if dirpath and os.path.exists(dirpath):
        shutil.rmtree(dirpath, ignore_errors=True)


# ============================================================
# 5. 页面 & 全局 API
# ============================================================
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/ping')
def api_ping():
    return jsonify({"status": "ok", "time": datetime.now().isoformat()})


@app.route('/api/config', methods=['GET'])
def api_get_config():
    return jsonify({"config": load_config(), "features": FEATURE_META})


@app.route('/api/config', methods=['POST'])
def api_save_config():
    try:
        data = request.get_json(force=True) or {}
        merged = load_config()
        for k in merged.keys():
            if k in data:
                merged[k] = bool(data[k])
        save_config(merged)
        return jsonify({"status": "ok", "config": merged})
    except Exception as e:
        return err(f"保存失败: {e}", 500)


@app.route('/api/download/<path:filename>')
def api_download(filename):
    safe = os.path.basename(filename)
    path = os.path.join(DOWNLOAD_FOLDER, safe)
    if not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=safe)


# ============================================================
# 6. 通用样式辅助
# ============================================================
def apply_excel_header_style(ws, fill_color='1F5BAA'):
    fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
    font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
    border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF'),
    )
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    ws.row_dimensions[1].height = 24


def apply_excel_body_style(ws):
    border = Border(
        left=Side(style='thin', color='DCDCDC'),
        right=Side(style='thin', color='DCDCDC'),
        top=Side(style='thin', color='DCDCDC'),
        bottom=Side(style='thin', color='DCDCDC'),
    )
    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(wrap_text=True, vertical='center')
            cell.border = border
            cell.font = Font(name='微软雅黑', size=10)


def set_chinese_font(run, font_name='宋体', size=12):
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    existing = rPr.find(qn('w:rFonts'))
    if existing is not None:
        rPr.remove(existing)
    rPr.append(rFonts)
    if size:
        run.font.size = Pt(size)


# ============================================================
# 7. 功能 1: Word 多文档一键合并
# ============================================================
@app.route('/api/word/merge', methods=['POST'])
def api_word_merge():
    if not is_enabled('word_merge'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files:
        return err("请至少上传1个docx文件", 400)

    work_dir, paths = save_uploads(files, 'word_merge')
    try:
        merged = Document()
        normal = merged.styles['Normal']
        normal.font.size = Pt(12)
        rPr = normal.element.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:eastAsia'), '宋体')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        existing = rPr.find(qn('w:rFonts'))
        if existing is not None:
            rPr.remove(existing)
        rPr.append(rFonts)

        for idx, p in enumerate(paths):
            sub = Document(p)
            for element in sub.element.body.iterchildren():
                if element.tag == qn('w:sectPr'):
                    continue
                merged.element.body.append(deepcopy(element))
            if idx < len(paths) - 1:
                page_p = OxmlElement('w:p')
                r = OxmlElement('w:r')
                br = OxmlElement('w:br')
                br.set(qn('w:type'), 'page')
                r.append(br)
                page_p.append(r)
                merged.element.body.append(page_p)

        out = gen_filename('合并文档', 'docx')
        merged.save(os.path.join(DOWNLOAD_FOLDER, out))
        return jsonify({
            "status": "ok", "filename": out,
            "download_url": f"/api/download/{out}",
            "merged_count": len(paths)
        })
    except Exception as e:
        return err(f"合并失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 8. 功能 2: 文档脱敏与排版
# ============================================================
PHONE_RE = re.compile(r'(?<!\d)1[3-9]\d{9}(?!\d)')
TEL_RE = re.compile(r'(?<!\d)0\d{2,3}[-\s]?\d{7,8}(?!\d)')
ID_RE = re.compile(r'(?<!\d)\d{17}[\dXx](?!\d)')

PHONE_MASK = '1XX**********'
ID_MASK = '******************'
TEL_MASK = '0XXX-XXXXXXX'


def _apply_first_line_indent(p_elem, indent_chars=200):
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p_elem.insert(0, pPr)
    for old in pPr.findall(qn('w:ind')):
        pPr.remove(old)
    ind = OxmlElement('w:ind')
    ind.set(qn('w:firstLineChars'), str(indent_chars))
    pPr.append(ind)


def _redact_paragraph(para):
    if not para.text.strip():
        return False
    original = para.text
    new_text = PHONE_RE.sub(PHONE_MASK, original)
    new_text = ID_RE.sub(ID_MASK, new_text)
    new_text = TEL_RE.sub(TEL_MASK, new_text)
    if new_text == original:
        return False
    for r in list(para.runs):
        r._element.getparent().remove(r._element)
    run = para.add_run(new_text)
    set_chinese_font(run, '宋体', None)
    return True


@app.route('/api/word/redact', methods=['POST'])
def api_word_redact():
    if not is_enabled('word_redact'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files:
        return err("请上传1个docx文件", 400)

    work_dir, paths = save_uploads(files, 'word_redact')
    try:
        doc = Document(paths[0])

        redact_count = 0
        for para in doc.paragraphs:
            if _redact_paragraph(para):
                redact_count += 1
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if _redact_paragraph(para):
                            redact_count += 1

        for para in doc.paragraphs:
            if para.text.strip():
                _apply_first_line_indent(para._element, 200)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para.text.strip():
                            _apply_first_line_indent(para._element, 200)

        out = gen_filename('脱敏排版', 'docx')
        doc.save(os.path.join(DOWNLOAD_FOLDER, out))
        return jsonify({
            "status": "ok", "filename": out,
            "download_url": f"/api/download/{out}",
            "redact_count": redact_count
        })
    except Exception as e:
        return err(f"脱敏失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 9. 功能 3: 关键信息提取库
# ============================================================
TIME_RE = re.compile(
    r'\d{4}[-年/.]\d{1,2}[-月/.]\d{1,2}[日号]?'
    r'(?:\s?\d{1,2}[:：]\d{1,2}(?::\d{1,2})?)?'
)
SHORTTIME_RE = re.compile(r'\b\d{1,2}[:：]\d{2}(?::\d{1,2})?\b')
MOBILE_RE = re.compile(r'(?<!\d)1[3-9]\d{9}(?!\d)')
LANDLINE_RE = re.compile(r'(?<!\d)0\d{2,3}[-\s]?\d{7,8}(?!\d)')
EMAIL_RE = re.compile(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+')
IDCARD_RE = re.compile(r'(?<!\d)\d{17}[\dXx](?!\d)')
ADDR_RE = re.compile(
    r'[\u4e00-\u9fa5]{2,}'
    r'(?:省|自治区|市|区|县|镇|乡|街道|路|街|道|村|巷|号|栋|单元|室)'
    r'[\u4e00-\u9fa5\d]{0,30}'
)
MONEY_RE = re.compile(
    r'(?:￥|¥|RMB|人民币)?\s?\d[\d,]*(?:\.\d+)?\s?[元万千百]'
)


def extract_text_from_docx(path):
    doc = Document(path)
    parts = []
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    parts.append(cell.text)
    return '\n'.join(parts)


@app.route('/api/word/extract', methods=['POST'])
def api_word_extract():
    if not is_enabled('word_extract'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files:
        return err("请至少上传1个docx文件", 400)

    work_dir, paths = save_uploads(files, 'word_extract')
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "关键信息汇总"
        headers = ['源文件', '时间', '手机号', '座机', '身份证号', '邮箱', '金额', '疑似地址', '正文长度']
        ws.append(headers)
        apply_excel_header_style(ws)

        for p in paths:
            text = extract_text_from_docx(p)
            times = sorted(set(TIME_RE.findall(text) + SHORTTIME_RE.findall(text)))
            mobiles = sorted(set(MOBILE_RE.findall(text)))
            tels = sorted(set(LANDLINE_RE.findall(text)))
            ids = sorted(set(IDCARD_RE.findall(text)))
            emails = sorted(set(EMAIL_RE.findall(text)))
            money = sorted(set(MONEY_RE.findall(text)))
            addrs = sorted(set(ADDR_RE.findall(text)))

            row = [
                os.path.basename(p).split('_', 1)[-1] if '_' in os.path.basename(p) else os.path.basename(p),
                ' / '.join(times[:8]),
                ' / '.join(mobiles),
                ' / '.join(tels),
                ' / '.join(ids),
                ' / '.join(emails),
                ' / '.join(money[:6]),
                ' / '.join(addrs[:3]),
                len(text)
            ]
            ws.append(row)

        apply_excel_body_style(ws)
        col_widths = {'A': 26, 'B': 28, 'C': 18, 'D': 16, 'E': 24, 'F': 24, 'G': 16, 'H': 30, 'I': 10}
        for col, w in col_widths.items():
            ws.column_dimensions[col].width = w
        ws.freeze_panes = 'A2'

        out = gen_filename('关键信息汇总', 'xlsx')
        wb.save(os.path.join(DOWNLOAD_FOLDER, out))
        return jsonify({
            "status": "ok", "filename": out,
            "download_url": f"/api/download/{out}",
            "file_count": len(paths)
        })
    except Exception as e:
        return err(f"提取失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 10. 功能 4: Excel 大表按分类拆分
# ============================================================
@app.route('/api/excel/split', methods=['POST'])
def api_excel_split():
    if not is_enabled('excel_split'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files:
        return err("请上传1个xlsx文件", 400)
    column = (request.form.get('column', '') or '部门').strip()

    work_dir, paths = save_uploads(files, 'excel_split')
    try:
        wb = load_workbook(paths[0], data_only=True)
        ws = wb.active
        headers = [c.value for c in ws[1]]
        if column not in headers:
            return err(f"未找到列 [{column}]，当前列名: {list(headers)}", 400)
        col_idx = headers.index(column)

        groups = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            key = row[col_idx] if col_idx < len(row) else None
            if key is None or str(key).strip() == '':
                key = '未分类'
            key = str(key).strip()
            groups.setdefault(key, []).append(row)

        zip_name = gen_filename('拆分结果', 'zip')
        zip_path = os.path.join(DOWNLOAD_FOLDER, zip_name)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for key, rows in groups.items():
                wb2 = Workbook()
                ws2 = wb2.active
                ws2.title = str(key)[:30]
                ws2.append(list(headers))
                for r in rows:
                    ws2.append(list(r))
                apply_excel_header_style(ws2)
                for i, h in enumerate(headers, 1):
                    col_letter = ws2.cell(row=1, column=i).column_letter
                    ws2.column_dimensions[col_letter].width = 18

                safe_key = re.sub(r'[\\/:*?"<>|\s]', '_', key)[:40] or '未分类'
                arcname = f"{column}_{safe_key}.xlsx"
                tmp_path = os.path.join(work_dir, arcname)
                wb2.save(tmp_path)
                zf.write(tmp_path, arcname=arcname)

        summary = {k: len(v) for k, v in groups.items()}
        return jsonify({
            "status": "ok", "filename": zip_name,
            "download_url": f"/api/download/{zip_name}",
            "summary": summary,
            "split_count": len(groups)
        })
    except Exception as e:
        return err(f"拆分失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 11. 功能 5: Excel 多表横向对齐合并
# ============================================================
@app.route('/api/excel/merge', methods=['POST'])
def api_excel_merge():
    if not is_enabled('excel_merge'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files or len(files) < 2:
        return err("请至少上传2个xlsx", 400)
    keycol = (request.form.get('keycol', '') or '身份证号').strip()

    work_dir, paths = save_uploads(files, 'excel_merge')
    try:
        main_wb = load_workbook(paths[0], data_only=True)
        main_ws = main_wb.active
        main_headers = [c.value for c in main_ws[1]]
        if keycol not in main_headers:
            return err(f"主表中未找到列 [{keycol}]，当前列名: {list(main_headers)}", 400)
        key_idx = main_headers.index(keycol)
        main_data = list(main_ws.iter_rows(min_row=2, values_only=True))

        added_cols = []
        new_col_values = []

        for path in paths[1:]:
            wb = load_workbook(path, data_only=True)
            ws = wb.active
            sub_headers = [c.value for c in ws[1]]
            if keycol not in sub_headers:
                continue
            sub_key_idx = sub_headers.index(keycol)
            sub_data = {}
            for r in ws.iter_rows(min_row=2, values_only=True):
                if sub_key_idx < len(r) and r[sub_key_idx] is not None:
                    sub_data[str(r[sub_key_idx]).strip()] = r

            base_name = os.path.splitext(os.path.basename(path))[0]
            base_name = re.sub(r'^\d+_', '', base_name)
            for hi, h in enumerate(sub_headers):
                if h is None or h == keycol or h in main_headers:
                    continue
                col_name = f"{h}({base_name})"
                added_cols.append(col_name)
                col_vals = []
                for r in main_data:
                    k = str(r[key_idx]).strip() if key_idx < len(r) and r[key_idx] is not None else ''
                    sub_row = sub_data.get(k)
                    if sub_row is not None and hi < len(sub_row):
                        col_vals.append(sub_row[hi])
                    else:
                        col_vals.append('')
                new_col_values.append(col_vals)

        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws.title = "横向合并结果"
        new_headers = list(main_headers) + added_cols
        new_ws.append(new_headers)
        for i, row in enumerate(main_data):
            new_row = list(row) + [col[i] for col in new_col_values]
            new_ws.append(new_row)

        apply_excel_header_style(new_ws)
        apply_excel_body_style(new_ws)
        for i, h in enumerate(new_headers, 1):
            col_letter = new_ws.cell(row=1, column=i).column_letter
            new_ws.column_dimensions[col_letter].width = 16
        new_ws.freeze_panes = 'A2'

        out = gen_filename('横向合并', 'xlsx')
        new_wb.save(os.path.join(DOWNLOAD_FOLDER, out))
        return jsonify({
            "status": "ok", "filename": out,
            "download_url": f"/api/download/{out}",
            "added_columns": added_cols
        })
    except Exception as e:
        return err(f"合并失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 12. 功能 6: Excel 身份证信息智能解析
# ============================================================
def _norm_id_cell(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    if 'e' in s.lower():
        try:
            s = f"{float(s):.0f}"
        except Exception:
            return s
    return s


def parse_idcard(num):
    s = _norm_id_cell(num)
    if not s or len(s) != 18:
        return ('', '', '')
    try:
        birth = f"{s[6:10]}-{s[10:12]}-{s[12:14]}"
        gender = '男' if int(s[16]) % 2 == 1 else '女'
        today = date.today()
        b_date = date(int(s[6:10]), int(s[10:12]), int(s[12:14]))
        age = today.year - b_date.year - ((today.month, today.day) < (b_date.month, b_date.day))
        return (birth, gender, age)
    except Exception:
        return ('', '', '')


@app.route('/api/excel/idcard', methods=['POST'])
def api_excel_idcard():
    if not is_enabled('excel_idcard'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files:
        return err("请上传1个xlsx文件", 400)
    idcol = (request.form.get('idcol', '') or '').strip()

    work_dir, paths = save_uploads(files, 'excel_idcard')
    try:
        wb = load_workbook(paths[0], data_only=True)
        ws = wb.active
        headers = [c.value for c in ws[1]]

        id_idx = None
        if idcol and idcol in headers:
            id_idx = headers.index(idcol)
        else:
            for i, h in enumerate(headers):
                if h and ('身份证' in str(h) or '证件号' in str(h) or 'ID' in str(h).upper()):
                    id_idx = i
                    break
        if id_idx is None:
            id_re_local = re.compile(r'^\d{17}[\dXx]$')
            for r in ws.iter_rows(min_row=2, max_row=min(ws.max_row, 30), values_only=True):
                for i, v in enumerate(r):
                    n = _norm_id_cell(v)
                    if n and id_re_local.match(n):
                        id_idx = i
                        break
                if id_idx is not None:
                    break
        if id_idx is None:
            return err("未识别到身份证号列，请手动指定列名", 400)

        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws.title = ws.title
        new_headers = list(headers) + ['出生日期', '性别', '年龄']
        new_ws.append(new_headers)

        hit = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            row_list = list(row)
            id_val = row_list[id_idx] if id_idx < len(row_list) else None
            birth, gender, age = parse_idcard(id_val)
            if birth:
                hit += 1
            row_list += [birth, gender, age]
            new_ws.append(row_list)

        apply_excel_header_style(new_ws)
        apply_excel_body_style(new_ws)
        for i, _ in enumerate(new_headers, 1):
            col_letter = new_ws.cell(row=1, column=i).column_letter
            new_ws.column_dimensions[col_letter].width = 16
        new_ws.freeze_panes = 'A2'

        out = gen_filename('身份证解析', 'xlsx')
        new_wb.save(os.path.join(DOWNLOAD_FOLDER, out))
        return jsonify({
            "status": "ok", "filename": out,
            "download_url": f"/api/download/{out}",
            "id_column": headers[id_idx],
            "parsed_count": hit
        })
    except Exception as e:
        return err(f"解析失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 13. 功能 7: PPT 汇报大纲一键提取
# ============================================================
@app.route('/api/ppt/outline', methods=['POST'])
def api_ppt_outline():
    if not is_enabled('ppt_outline'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files:
        return err("请上传1个pptx文件", 400)

    work_dir, paths = save_uploads(files, 'ppt_outline')
    try:
        prs = Presentation(paths[0])
        doc = Document()
        style = doc.styles['Normal']
        style.font.size = Pt(11)
        rPr = style.element.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:eastAsia'), '微软雅黑')
        rFonts.set(qn('w:ascii'), 'Calibri')
        rFonts.set(qn('w:hAnsi'), 'Calibri')
        existing = rPr.find(qn('w:rFonts'))
        if existing is not None:
            rPr.remove(existing)
        rPr.append(rFonts)

        title_para = doc.add_heading(f"PPT 汇报大纲", level=0)
        sub = doc.add_paragraph()
        sub_run = sub.add_run(f"源文件: {os.path.basename(paths[0])}    页数: {len(prs.slides)}")
        sub_run.italic = True
        sub_run.font.size = Pt(10)
        sub_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

        for i, slide in enumerate(prs.slides, 1):
            title_text = ''
            if slide.shapes.title and slide.shapes.title.has_text_frame:
                title_text = slide.shapes.title.text_frame.text.strip()
            if not title_text:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        t = shape.text_frame.text.strip()
                        if t:
                            title_text = t.split('\n')[0]
                            break
            doc.add_heading(f"第 {i} 页: {title_text or '（无标题）'}", level=1)

            for shape in slide.shapes:
                if shape == slide.shapes.title:
                    continue
                if shape.has_text_frame:
                    txt = shape.text_frame.text.strip()
                    if txt and txt != title_text:
                        for line in txt.split('\n'):
                            line = line.strip()
                            if line:
                                p = doc.add_paragraph(line, style='List Bullet')
                elif shape.has_table:
                    doc.add_paragraph("表格内容:", style='Intense Quote')
                    for r in shape.table.rows:
                        doc.add_paragraph(' | '.join(c.text.strip() for c in r.cells))

        out = gen_filename('PPT大纲', 'docx')
        doc.save(os.path.join(DOWNLOAD_FOLDER, out))
        return jsonify({
            "status": "ok", "filename": out,
            "download_url": f"/api/download/{out}",
            "slide_count": len(prs.slides)
        })
    except Exception as e:
        return err(f"提取失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 14. 功能 8: PPT 批量加防伪水印
# ============================================================
def add_watermark_to_slide(slide, text, sw, sh):
    box_w = int(sw * 0.6)
    box_h = int(sh * 0.3)
    left = (sw - box_w) // 2
    top = (sh - box_h) // 2
    txBox = slide.shapes.add_textbox(left, top, box_w, box_h)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = PptxPt(54)
    run.font.bold = True
    run.font.color.rgb = PptxRGBColor(0xCC, 0xCC, 0xCC)
    sp = txBox._element
    if sp.spPr is not None and sp.spPr.xfrm is not None:
        sp.spPr.xfrm.set('rot', str(-30 * 60000))


@app.route('/api/ppt/watermark', methods=['POST'])
def api_ppt_watermark():
    if not is_enabled('ppt_watermark'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files:
        return err("请至少上传1个pptx", 400)
    wm_text = (request.form.get('text', '') or '内部资料 严禁外传').strip()

    work_dir, paths = save_uploads(files, 'ppt_watermark')
    out_files = []
    try:
        for path in paths:
            prs = Presentation(path)
            sw, sh = prs.slide_width, prs.slide_height
            for slide in prs.slides:
                add_watermark_to_slide(slide, wm_text, sw, sh)
            out = gen_filename('水印版', 'pptx')
            prs.save(os.path.join(DOWNLOAD_FOLDER, out))
            out_files.append(out)

        if len(out_files) == 1:
            return jsonify({
                "status": "ok", "filename": out_files[0],
                "download_url": f"/api/download/{out_files[0]}",
                "watermark": wm_text
            })
        else:
            zip_name = gen_filename('水印打包', 'zip')
            zip_path = os.path.join(DOWNLOAD_FOLDER, zip_name)
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for f in out_files:
                    zf.write(os.path.join(DOWNLOAD_FOLDER, f), arcname=f)
            return jsonify({
                "status": "ok", "filename": zip_name,
                "download_url": f"/api/download/{zip_name}",
                "watermark": wm_text
            })
    except Exception as e:
        return err(f"水印失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 15. 功能 9: PPT 图片提取打包
# ============================================================
@app.route('/api/ppt/images', methods=['POST'])
def api_ppt_images():
    if not is_enabled('ppt_images'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files:
        return err("请至少上传1个pptx", 400)

    work_dir, paths = save_uploads(files, 'ppt_images')
    try:
        zip_name = gen_filename('PPT图片', 'zip')
        zip_path = os.path.join(DOWNLOAD_FOLDER, zip_name)
        total = 0
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for path in paths:
                base = os.path.splitext(os.path.basename(path))[0]
                base = re.sub(r'^\d+_', '', base)
                with zipfile.ZipFile(path, 'r') as src:
                    for name in src.namelist():
                        if name.startswith('ppt/media/') and not name.endswith('/'):
                            data = src.read(name)
                            ext = os.path.splitext(name)[1]
                            local = f"{base}_{os.path.basename(name)}"
                            zf.writestr(local, data)
                            total += 1
        return jsonify({
            "status": "ok", "filename": zip_name,
            "download_url": f"/api/download/{zip_name}",
            "image_count": total
        })
    except Exception as e:
        return err(f"提取失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 16. 功能 10: PDF 一键合并
# ============================================================
@app.route('/api/pdf/merge', methods=['POST'])
def api_pdf_merge():
    if not is_enabled('pdf_merge'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files or len(files) < 2:
        return err("请至少上传2个pdf文件", 400)

    work_dir, paths = save_uploads(files, 'pdf_merge')
    try:
        merger = PdfMerger()
        for p in paths:
            merger.append(p)
        out = gen_filename('合并PDF', 'pdf')
        merger.write(os.path.join(DOWNLOAD_FOLDER, out))
        merger.close()
        return jsonify({
            "status": "ok", "filename": out,
            "download_url": f"/api/download/{out}",
            "merged_count": len(paths)
        })
    except Exception as e:
        return err(f"合并失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 17. 功能 11: 批量文件规范化重命名
# ============================================================
@app.route('/api/file/rename', methods=['POST'])
def api_file_rename():
    if not is_enabled('file_rename'):
        return err("该功能已被管理员关闭", 403)

    files = request.files.getlist('files')
    if not files:
        return err("请上传文件", 400)
    prefix = request.form.get('prefix', '') or ''
    suffix = request.form.get('suffix', '') or ''

    work_dir, paths = save_uploads(files, 'file_rename')
    try:
        zip_name = gen_filename('重命名结果', 'zip')
        zip_path = os.path.join(DOWNLOAD_FOLDER, zip_name)
        renamed = []
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for p in paths:
                orig_name = os.path.basename(p)
                orig_name = re.sub(r'^\d+_', '', orig_name)
                stem, ext = os.path.splitext(orig_name)
                new_name = f"{prefix}{stem}{suffix}{ext}"
                new_name = re.sub(r'[\\/:*?"<>|]', '_', new_name)
                zf.write(p, arcname=new_name)
                renamed.append({"orig": orig_name, "new": new_name})

        return jsonify({
            "status": "ok", "filename": zip_name,
            "download_url": f"/api/download/{zip_name}",
            "renamed": renamed
        })
    except Exception as e:
        return err(f"重命名失败: {e}", 500)
    finally:
        safe_clear(work_dir)


# ============================================================
# 18. 功能 12: 长文本极速清洗
# ============================================================
CONTROL_RE = re.compile(r'[\x00-\x08\x0b-\x1f\x7f]')
ZERO_WIDTH_RE = re.compile(r'[\u200b-\u200f\u2028-\u202f\ufeff\uFFFD]')


def clean_text(text):
    text = CONTROL_RE.sub('', text)
    text = ZERO_WIDTH_RE.sub('', text)
    lines = [l.strip() for l in text.split('\n')]
    lines = [l for l in lines if l]
    return '\n'.join(lines)


@app.route('/api/text/clean', methods=['POST'])
def api_text_clean():
    if not is_enabled('text_clean'):
        return err("该功能已被管理员关闭", 403)

    keyword = (request.form.get('keyword', '') or '').strip()
    text = request.form.get('text', '') or ''

    files = request.files.getlist('files')
    if files:
        work_dir, paths = save_uploads(files, 'text_clean')
    else:
        work_dir = None
        paths = []

    try:
        if not text and paths:
            with open(paths[0], 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
        if not text:
            return err("请输入文本或上传txt文件", 400)

        original_len = len(text)
        cleaned = clean_text(text)
        cleaned_len = len(cleaned)

        if keyword:
            sentences = re.split(r'(?<=[。！？!?\n])', cleaned)
            matched = [s.strip() for s in sentences if keyword in s and s.strip()]
            result = '\n'.join(matched)
            stat = f"原文 {original_len} 字 / 清洗后 {cleaned_len} 字 / 命中关键词 [{keyword}] 的句子 {len(matched)} 句"
        else:
            result = cleaned
            stat = f"原文 {original_len} 字 / 清洗后 {cleaned_len} 字"

        out = gen_filename('文本清洗结果', 'txt')
        with open(os.path.join(DOWNLOAD_FOLDER, out), 'w', encoding='utf-8') as f:
            f.write("=" * 60 + "\n")
            f.write("内网文档处理工作台 — 长文本极速清洗报告\n")
            f.write("=" * 60 + "\n")
            f.write(f"统计: {stat}\n")
            f.write(f"关键词: {keyword or '（未指定，输出全部）'}\n")
            f.write("=" * 60 + "\n\n")
            f.write(result)

        return jsonify({
            "status": "ok", "filename": out,
            "download_url": f"/api/download/{out}",
            "stat": stat
        })
    except Exception as e:
        return err(f"清洗失败: {e}", 500)
    finally:
        if work_dir:
            safe_clear(work_dir)


# ============================================================
# 19. 错误处理
# ============================================================
@app.errorhandler(413)
def too_large(e):
    return jsonify({"status": "error", "message": "上传文件超过 500MB 上限"}), 413


@app.errorhandler(404)
def not_found(e):
    if request.path.startswith('/api/'):
        return jsonify({"status": "error", "message": "资源不存在"}), 404
    return render_template('index.html')


@app.errorhandler(500)
def server_error(e):
    if request.path.startswith('/api/'):
        return jsonify({"status": "error", "message": "服务内部错误，请查看控制台日志"}), 500
    return render_template('index.html')


# ============================================================
# 20. 启动
# ============================================================
DEFAULT_HOST = '127.0.0.1'
DEFAULT_PORT = 5000
PORT_SCAN_LIMIT = 12


def _port_available(host, port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        return s.connect_ex((host, port)) != 0


def _pick_port(host, start, limit=PORT_SCAN_LIMIT):
    for port in range(start, start + limit):
        if _port_available(host, port):
            return port
    return start


def _print_banner(url, port_shifted):
    line = '=' * 64
    print(line)
    print(f"  {APP_NAME}  {APP_VERSION}")
    print(line)
    print(f"  访问地址   {url}")
    print(f"  程序目录   {DATA_DIR}")
    print(f"  结果文件   {DOWNLOAD_FOLDER}")
    print(f"  配置文件   {CONFIG_FILE}")
    print(line)
    if port_shifted:
        print(f"  注意: 默认端口 {DEFAULT_PORT} 已被占用，已自动切换到 {url.rsplit(':', 1)[-1]} 端口。")
    print("  运行期间请保持本窗口开启；关闭窗口即停止服务。")
    print("  浏览器若未自动打开，请手动访问上方地址。")
    print()
    sys.stdout.flush()


def main():
    host = os.environ.get('WORKBENCH_HOST', DEFAULT_HOST).strip() or DEFAULT_HOST
    try:
        start_port = int(os.environ.get('WORKBENCH_PORT', DEFAULT_PORT))
    except (TypeError, ValueError):
        start_port = DEFAULT_PORT

    port = _pick_port(host, start_port)
    url = f"http://{host}:{port}"

    logging.getLogger('werkzeug').setLevel(logging.ERROR)

    _print_banner(url, port != start_port)

    if os.environ.get('WORKBENCH_NO_BROWSER') != '1':
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()

    run_simple(host, port, app, use_reloader=False, use_debugger=False, threaded=True)


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n服务已停止。")
    except OSError as e:
        print(f"\n启动失败: {e}")
        print("请确认端口未被占用，或通过环境变量 WORKBENCH_PORT 指定其他端口。")
        if _is_frozen():
            input("\n按回车键退出...")
    except Exception:
        import traceback
        traceback.print_exc()
        if _is_frozen():
            input("\n程序异常退出，按回车键关闭窗口...")
        sys.exit(1)

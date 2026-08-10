# -*- coding: utf-8 -*-
"""UNO 修复: 刀锋文档长数字串折行对齐 Word
原理: LO 把"数字+顿号"长串当作一个不可拆西文单词(不在 、 处断行),
     而 Word 按有效字宽(全角9.2pt/数字4.92pt @10pt)逐字贪婪填充。
     本脚本模拟 Word 折行, 在断行位置插入显式 LINE_BREAK, 强制 LO 对齐。
仅处理含 >=60 字符长数字串的段落(刀锋类), 其他段落不动。
用法: /opt/libreoffice7.6/program/python uno_linewrap_fix.py <输入URL> <输出pdf URL>
"""
import re
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

RUN_RE = re.compile(u'[0-9\u3001\uFF0C\u3002,.\uFF0E]{60,}')

# Word 自然字宽(未压缩): ASCII(数字/半角) = 0.5em, 全角 = 1.0em
# 断行规则(从两份基准 PDF 逆向): 数字串 + 其后全角标点 = 原子 token,
# 断点只允许在 token 末尾; 极限容量 = 样式文本宽 + 4.6pt(Word 比 LO 宽) + 14.1pt(标点悬挂)
HANG_PT = 14.1


def P(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect():
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)
    for _ in range(30):
        try:
            return resolver.resolve(
                "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
        except Exception:
            time.sleep(1)
    raise RuntimeError("cannot connect")


def tokenize(text, em):
    """切成 (宽度, 末尾可断, 结束下标) 序列。
    数字/半角逗号点 连串 + 可选的一个全角标点 = 原子 token(只在末尾可断);
    CJK 逐字可断; 全角闭合标点并入前项(禁止行首)。"""
    W_ASCII = 0.5 * em
    W_FULL = 1.0 * em
    GLUE_TAIL = u'\u3001\u3002\uFF0C\uFF1B\uFF1A'  # 、。，；：
    CLOSING = u'\u3001\u3002\uFF0C\uFF1B\uFF1A\uFF01\uFF1F\uFF09\u3011'
    items = []
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch.isdigit() or ch in ',.':
            j = i
            while j < n and (text[j].isdigit() or text[j] in ',.'):
                j += 1
            w = (j - i) * W_ASCII
            if j < n and text[j] in GLUE_TAIL:
                w += W_FULL
                j += 1
                items.append([w, True, j])
            else:
                items.append([w, False, j])
            i = j
        elif ch in CLOSING and items:
            items[-1][0] += W_FULL
            items[-1][1] = True
            items[-1][2] = i + 1
            i += 1
        else:
            items.append([W_FULL, True, i + 1])
            i += 1
    return items


def simulate_breaks_token(text, limit_pt, em):
    """含"数字,数字"胶合的文档: 按 token 断行(老多页变体实测校准)"""
    items = tokenize(text, em)
    breaks = []
    x = 0.0
    last_ok_end = None
    x_at_last_ok = 0.0
    for w, brk, end in items:
        if x + w > limit_pt and x > 0:
            if last_ok_end is not None:
                breaks.append(last_ok_end)
                x = x - x_at_last_ok + w
            else:
                breaks.append(end - 1)
                x = w
        else:
            x += w
        if brk:
            last_ok_end = end
            x_at_last_ok = x
    return [b for b in breaks if b > 0]


def simulate_breaks_compressed(text, limit_pt, em):
    """全角标点数字串文档: 按压缩字宽逐字贪婪(mp1 变体实测校准)"""
    breaks = []
    x = 0.0
    for i, ch in enumerate(text):
        if ch in u'\r\n\x0b\x0c':
            x = 0.0
            continue
        w = (0.492 if ord(ch) < 0x80 else 0.92) * em
        if x + w > limit_pt and x > 0:
            breaks.append(i)
            x = w
        else:
            x += w
    return breaks


def get_text_width_pt(doc):
    """从页面样式取正文宽度(pt)"""
    style_families = doc.StyleFamilies.getByName("PageStyles")
    names = style_families.getElementNames()
    st = style_families.getByName(names[0])
    w_100mm = st.Width - st.LeftMargin - st.RightMargin
    return w_100mm * 72.0 / 2540.0


# Word 有效文本宽度比 LO 页面样式值宽约 4.6pt(实测校准);
# 负右缩进让 LO 段内容得下模拟出的宽行(宽到 440pt, 留余量防 LO 重排)
DELTA_100MM = 840  # ≈ 23.8pt


def process_text(text_obj, doc, limit_pt):
    fixed = 0
    paras = text_obj.createEnumeration()
    while paras.hasMoreElements():
        para = paras.nextElement()
        if not para.supportsService("com.sun.star.text.Paragraph"):
            continue
        s = para.getString()
        if not RUN_RE.search(s):
            continue
        em = 10.0
        try:
            em = para.CharHeight or 10.0
        except Exception:
            pass
        try:
            para.ParaRightMargin = para.ParaRightMargin - DELTA_100MM
        except Exception:
            pass
        if re.search(r'[0-9],[0-9]', s):
            # 含"数字,数字"胶合 → token 模型 (limit 含标点悬挂余量)
            limit = limit_pt + (4.6 + HANG_PT) * (em / 10.0)
            breaks = simulate_breaks_token(s, limit, em)
        else:
            # 全角标点数字串 → 压缩逐字模型
            limit = limit_pt + 4.6 * (em / 10.0)
            breaks = simulate_breaks_compressed(s, limit, em)
        for idx in reversed(breaks):
            cur = text_obj.createTextCursorByRange(para.getStart())
            cur.goRight(idx, False)
            text_obj.insertControlCharacter(cur, 1, False)  # 1 = LINE_BREAK
            fixed += 1
    return fixed


def walk_tables(text_obj, doc, limit_pt):
    n = 0
    try:
        tables = text_obj.getTextTables()
    except Exception:
        return 0
    for i in range(tables.getCount()):
        tbl = tables.getByIndex(i)
        for name in tbl.getCellNames():
            cell = tbl.getCellByName(name)
            n += process_text(cell, doc, limit_pt)
            n += walk_tables(cell, doc, limit_pt)
    return n


def main():
    src_url, dst_url = sys.argv[1], sys.argv[2]
    ctx = connect()
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.loadComponentFromURL(src_url, "_blank", 0, (P("Hidden", True),))
    try:
        limit_pt = get_text_width_pt(doc)
        print("text_width_pt: %.2f" % limit_pt)
        n = process_text(doc.Text, doc, limit_pt)
        n += walk_tables(doc.Text, doc, limit_pt)
        print("line_breaks_inserted:", n)
        doc.storeToURL(dst_url, (P("FilterName", "writer_pdf_Export"),))
        print("exported:", dst_url)
    finally:
        doc.close(False)


if __name__ == '__main__':
    main()

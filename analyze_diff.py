#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
对比原始版(V3对外公示)和修改稿(temp.docx)，
以原始版为基础，将润色修改过的文字段落标蓝色，
标题/小标题段落不标蓝，保留原始黑色。
"""
import zipfile
from lxml import etree
import os, re, difflib

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

SRC = '/home/admin/lindy-growth/original.docx'
DST = '/home/admin/lindy-growth/\u53ee\u549a\u4e70\u83dc\u6848\u4f8b_\u4fee\u6539\u7a3f_\u5e26\u6807\u6ce8.docx'

# ---------------------------
# Paragraph mapping: original vs modified
# Based on the comparison above
# ---------------------------

# Original paragraphs (0-indexed in original doc)
# Modified paragraphs (0-indexed in temp.docx)
# We need to map which original paragraph maps to which modified paragraph

# Key: original paragraph index that was MODIFIED (text changed but same content)
# These are paragraphs where the text was polished/rewritten
# Title paragraphs should NOT be marked

# Build text map
with zipfile.ZipFile(SRC, 'r') as z:
    orig_xml = z.read('word/document.xml')
orig_root = etree.fromstring(orig_xml)

with zipfile.ZipFile('/home/admin/lindy-growth/temp.docx', 'r') as z:
    mod_xml = z.read('word/document.xml')
mod_root = etree.fromstring(mod_xml)

def get_para_text(para):
    texts = []
    for t in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)

orig_paras = list(orig_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'))
mod_paras = list(mod_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'))

# Build mapping: for each original para, find the corresponding modified para
# by matching content
orig_idx_to_mod_idx = {}
mod_used = set()

for oi, op in enumerate(orig_paras):
    ot = get_para_text(op).strip()
    if not ot:
        continue
    for mi, mp in enumerate(mod_paras):
        if mi in mod_used:
            continue
        mt = get_para_text(mp).strip()
        # Check if they are related (same topic)
        if ot == mt or ot[:10] == mt[:10] or (len(ot) > 30 and ot[:30] == mt[:30]):
            orig_idx_to_mod_idx[oi] = mi
            mod_used.add(mi)
            break

print("Original->Modified mapping:")
for oi, mi in sorted(orig_idx_to_mod_idx.items()):
    ot = get_para_text(orig_paras[oi]).strip()[:80]
    mt = get_para_text(mod_paras[mi]).strip()[:80]
    same = ot == mt
    print(f"  P{oi:2d} -> P{mi:2d} | {'SAME' if same else 'DIFF'} | {ot}")

# ---------------------------
# Find paragraphs that were modified (text changed)
# Exclude title/section-header paragraphs
# ---------------------------

# Title patterns: short text, ends with colon or is a standalone title
def is_title(text):
    t = text.strip()
    if not t:
        return False
    # Known titles
    titles = [
        '案例背景', '实施过程与路径', '塔尖：战略顶层化', '塔身：载体创新化',
        '塔基：生态协同化', '成果和价值', '资源保障', '高层示范',
        '栏目化内容运营', '智能化服务功能', '文化活动组织', '食安技能大赛',
        '案例一：', '上海壹佰米网络科技有限公司',
    ]
    clean = t.replace('\u201c', '').replace('\u201d', '').replace('\uff1a', ':')
    if t in titles:
        return True
    if t.endswith('：') or t.endswith(':'):
        return True
    if len(t) < 15 and '：' in t:
        return True
    return False

# Which paragraphs have different text
modified_paras = set()
for oi, mi in orig_idx_to_mod_idx.items():
    ot = get_para_text(orig_paras[oi]).strip()
    mt = get_para_text(mod_paras[mi]).strip()
    if ot != mt and not is_title(ot) and not is_title(mt):
        modified_paras.add(mi)

print(f"\nModified (non-title) paragraphs: {sorted(modified_paras)}")

# Also find NEW paragraphs in modified doc that don't map to any original
all_mod_mapped = set(orig_idx_to_mod_idx.values())
new_paras = set()
for mi, mp in enumerate(mod_paras):
    mt = get_para_text(mp).strip()
    if mt and mi not in all_mod_mapped:
        new_paras.add(mi)

print(f"New paragraphs (no original match): {sorted(new_paras)}")

# ---------------------------
# Now process the OUTPUT based on ORIGINAL
# We use original.docx as base, modify its text and mark changes in blue
# ---------------------------

# Read original document
with zipfile.ZipFile(SRC, 'r') as z:
    orig_all = {f.filename: z.read(f.filename) for f in z.infolist()}

orig_xml = orig_all['word/document.xml']
root = etree.fromstring(orig_xml)

paras = list(root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'))

# Text replacements: for each modified original paragraph, use the polished text from mod doc
replacements = {}
for oi, mi in orig_idx_to_mod_idx.items():
    if oi in modified_paras or mi in modified_paras:
        ot = get_para_text(paras[oi]).strip()
        mt_text = get_para_text(mod_paras[mi]).strip()
        if ot != mt_text and not is_title(ot):
            replacements[oi] = mt_text
            print(f"  Replace P{oi}: '{ot[:50]}...' -> '{mt_text[:50]}...'")

# Also add any NEW content from modified doc that doesn't exist in original
# These new paragraphs are from the CCFA annotation and transition sentences
# We'll add them after their logical position
# P3 in original is "案例背景", in modified P3 is the CCFA annotation
# The CCFA annotation is ADDITIONAL content not in original

# Actually, the structure is shifted in modified doc.
# Let me trace exact mapping more carefully

print("\n--- Detailed comparison ---")
for oi, mi in sorted(orig_idx_to_mod_idx.items()):
    ot = get_para_text(orig_paras[oi]).strip()
    mt = get_para_text(mod_paras[mi]).strip()
    if ot != mt:
        print(f"\nP{oi}(orig) -> P{mi}(mod)")
        print(f"  ORIG: {ot[:120]}")
        print(f"  MOD:  {mt[:120]}")

PYEOF
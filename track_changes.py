#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate a Track Changes (revised mode) Word document.
Based on original.docx, apply polished text as tracked changes.
"""
import zipfile
from lxml import etree
from copy import deepcopy
import os, datetime, uuid

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

SRC = '/home/admin/lindy-growth/original.docx'
DST = '/home/admin/lindy-growth/\u53ee\u549a\u4e70\u83dc_\u82b1\u8138\u7a3f.docx'

# Track changes metadata
REV_AUTHOR = 'Lindy'
REV_DATE = '2026-05-08T02:00:00Z'
REV_ID_BASE = 1000

def get_para_text(para):
    texts = []
    for t in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)

def is_title_para(text):
    t = text.strip()
    if not t:
        return True
    titles = [
        '案例背景', '实施过程与路径', '塔尖：战略顶层化', '塔身：载体创新化',
        '塔基：生态协同化', '成果和价值', '资源保障', '高层示范',
        '栏目化内容运营', '智能化服务功能', '文化活动组织', '食安技能大赛',
        '案例一：', '上海壹佰米网络科技有限公司',
    ]
    if t in titles:
        return True
    if t.endswith('\uff1a') or t.endswith(':'):
        return True
    if len(t) < 15 and ('\uff1a' in t or '：' in t):
        return True
    return False

def make_revision_rpr(author=REV_AUTHOR, date=REV_DATE, rev_id=1):
    """Create revision rPr element for ins/del."""
    rpr = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    rpr_author = etree.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rAuthor')
    rpr_author.text = author
    rpr_date = etree.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rDate')
    rpr_date.text = date
    rpr_id = etree.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rId')
    rpr_id.text = str(rev_id)
    return rpr

def make_run_for_text(text, color_val=None, preserve_space=True):
    """Create a w:r element with text."""
    run = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    # Copy formatting from the parent run if available (we'll handle later)
    t_elem = etree.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    if preserve_space:
        t_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space', 'preserve')
    t_elem.text = text
    return run

# Read original
with zipfile.ZipFile(SRC, 'r') as z:
    xml_bytes = z.read('word/document.xml')

root = etree.fromstring(xml_bytes)
paras = list(root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'))

# ===== STEP 1: Identify which paragraphs to modify =====
# Same set as before - body paragraphs that need polishing
replacements = {}

# P4: 案例背景（在原始版中是第4段，0-indexed）
replacements[4] = (
    '生鲜零售企业天然面临生鲜产品易腐易损、供应链条长、管理环节多的行业特性。'
    '在当前的消费环境下，用户需求已从\u201c价格敏感\u201d转向\u201c品质与安全敏感\u201d；'
    '与此同时，《食品安全法》《电子商务法》及97号令《企业落实食品安全主体责任监督管理规定》'
    '等政策法规持续趋严，不断强化企业主体责任；'
    '加之行业同质化竞争日益加剧，零售企业正面临严峻的食品安全挑战与运营压力。'
)

replacements[5] = (
    '叮咚买菜作为生鲜零售企业，自创立之初即将食品安全定位为\u201c企业生命线\u201d，'
    '提出\u201c绝不把不好的菜卖给用户\u201d的核心价值观，'
    '突破传统\u201c事后监管\u201d模式，确立\u201c文化先行、预防为主\u201d的治理理念，'
    '将食品安全文化建设提升至战略高度。'
    '企业将食品安全意识系统性地融入采购、仓储、加工、配送等全链条环节，'
    '围绕\u201c打造全员品质文化，实施全链条\u20187+1\u201d品控体系\u201d的质量方针持续推进，'
    '目标在于构建覆盖\u201c从田间到餐桌\u201d的全员参与、全链路渗透的食品安全文化体系，'
    '推动食安管理从\u201c被动合规\u201d向\u201c主动引领\u201d转型，将文化软实力转化为品质硬标准。'
)

replacements[7] = (
    '食品安全文化建设是一项意识渗透工程，既需\u201c自上而下\u201d的意识宣导，'
    '也需\u201c自下而上\u201d的行动推进。为实现食品安全行为的全面一致性，'
    '叮咚买菜通过\u201c金字塔式\u201d建设路径系统性地推广食品安全文化。'
)

replacements[9] = (
    '自创立以来，叮咚买菜创始人梁昌霖即提出\u201c绝不把不好的菜卖给用户\u201d的核心价值观。'
    '2025年，企业进一步提出\u201c不做所有人的75分，只做少数人的120分\u201d的战略定位，'
    '主动放弃行业内多数企业竞相追逐的低价策略与低价商品路线，'
    '转而以品质商品服务于对功能、品质、稀缺性等方面产品升级仍有明确需求的消费群体。'
)

replacements[11] = (
    '在食品安全投入方面，叮咚买菜品控中心配置近400名品控人员，'
    '覆盖源头、大仓、前置仓、自营工厂等全链路各环节。'
    '企业坚持到货批批验收检查，年验收批次近250万，品质不达标者坚决拒收；'
    '针对生鲜农产品安全问题，自建实验室并引入第三方权威机构驻仓，'
    '严格实施批批快检，年快检批次达百万级，药残不合格者坚决拒收；'
    '每月主动对市售时令类、高风险商品委托第三方机构随机定量抽检，'
    '抽检量逐年上升，累计抽检万余批次，检测不合格者坚决主动下架；'
    '对源头高风险工厂或基地实行100%现场审核通过方可准入，'
    '合作期间不定期实施非通知飞检，年开展第三方审核千余次。'
)

replacements[13] = (
    '叮咚买菜高级管理层全体亲身践行\u201c绝不把不好的菜卖给用户\u201d的核心价值观。'
    '除授予品控中心绝对的\u201c一票否决\u201d权之外，'
    '企业于每年供应商大会设立\u201c品质食力\u00b7年度品质杰出供应商\u201d奖项，'
    '由品控中心全权提名确认，重点表彰品质表现优异的供应商；'
    '每年夏初，高管团队拍摄夏季食安专项行动宣导片，预热并强调夏季食品安全风险，'
    '与公司全员同步各品类商品组的夏季食安行动计划。'
    '在任何关键季节、关键节点及公司策略调整之际，高级管理层始终秉持对品质的敬畏之心，'
    '在品质之路上坚持\u201c一寸窄，一公里深\u201d的经营方向。'
)

replacements[17] = (
    '自2018年起，叮咚买菜品控中心便创立内部品控订阅号【品控频道】，'
    '由品控专人运营管理，每周定时推送品控团队自行撰写的推文，'
    '将其作为品控中心与叮咚全员沟通的重要渠道，'
    '致力于在叮咚买菜内部宣传和打造全员品质文化，'
    '构建全员参与的食品安全文化阵地。'
)

replacements[19] = (
    '在频道内容方面，设置多重差异化内容方向，鼓励全员投稿，'
    '积极传播食品安全知识与行动成果。'
)

replacements[25] = (
    '在栏目功能方面，设置三大模块，涵盖员工食安知识库、内部建议反馈及全员排查反馈入口。'
)

replacements[32] = (
    '构建全员参与、全链路渗透的食品安全文化体系，'
    '需要借助多元化的形式与方法，提升员工的积极性与参与感，强化文化浸润效果。'
)

replacements[34] = (
    '通过组织丰富多样的食品安全文化活动，以高级管理层为示范，'
    '引导全员参与，并同步推进相应的食品安全行动。'
)

replacements[41] = (
    '以竞赛形式营造争优争先的团队氛围，为企业长远发展储备专业人才。'
    '技能大赛不局限于品控部门发起，更鼓励业务部门主动组织，'
    '以此促进人员专业技能的提升，激发员工潜能与质量意识。'
)

# ===== STEP 2: Apply Track Changes =====
rev_id_counter = [REV_ID_BASE]

for idx, new_text in replacements.items():
    para = paras[idx]
    orig_text = get_para_text(para)
    
    if is_title_para(orig_text):
        print(f"Skip P{idx} (title)")
        continue
    
    rev_id_counter[0] += 1
    rev_id = rev_id_counter[0]
    
    print(f"\nP{idx}: Applying tracked change (rev_id={rev_id})")
    print(f"  Old: {orig_text[:80]}...")
    print(f"  New: {new_text[:80]}...")
    
    # Get all existing runs in this paragraph to preserve their formatting
    existing_runs = list(para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'))
    
    # Build old text as del runs
    del_runs_parent = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del')
    rev_rpr = make_revision_rpr(rev_id=rev_id)
    del_runs_parent.append(rev_rpr)
    
    # For each existing run, create a w:del run
    for run in existing_runs:
        # Get the text and any formatting from this run
        run_texts = []
        run_rpr = None
        for child in run:
            tag = child.tag
            if tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t':
                if child.text:
                    run_texts.append(child.text)
            elif tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr':
                run_rpr = deepcopy(child)
        
        combined_text = ''.join(run_texts)
        if not combined_text:
            continue
        
        # Create w:delText element wrapped in w:r within w:del
        del_run = etree.SubElement(del_runs_parent, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        if run_rpr is not None:
            del_run.append(deepcopy(run_rpr))
        del_text = etree.SubElement(del_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}delText')
        del_text.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space', 'preserve')
        del_text.text = combined_text
    
    # Build new text as ins runs
    ins_runs_parent = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins')
    rev_rpr2 = make_revision_rpr(rev_id=rev_id)
    ins_runs_parent.append(rev_rpr2)
    
    # Split new text into runs (one run for simplicity)
    ins_run = etree.SubElement(ins_runs_parent, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    ins_t = etree.SubElement(ins_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    ins_t.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space', 'preserve')
    ins_t.text = new_text
    
    # Remove all existing runs from paragraph
    for run in existing_runs:
        para.remove(run)
    
    # Add del + ins to paragraph
    para.append(del_runs_parent)
    para.append(ins_runs_parent)
    
    print(f"  Done. Runs removed, del+ins added.")
    print(f"  Para children now: {len(list(para))}")

# ===== STEP 3: Write back to DOCX =====
new_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

# Copy all original files, replace document.xml
# But we also need to update the [Content_Types].xml to recognize revision mode
# Actually Track Changes doesn't need special content types - it's native OOXML

tmp = DST + '.tmp'
with zipfile.ZipFile(SRC, 'r') as zin:
    with zipfile.ZipFile(tmp, 'w') as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml':
                data = new_xml
            zout.writestr(item, data)

os.replace(tmp, DST)

# ===== STEP 4: Verify =====
with zipfile.ZipFile(DST, 'r') as z:
    check_xml = z.read('word/document.xml')
    check_root = etree.fromstring(check_xml)
    
    del_count = len(list(check_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del')))
    ins_count = len(list(check_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins')))
    
    print(f"\n{'='*50}")
    print(f"Track Changes Summary:")
    print(f"  Delete sections (w:del): {del_count}")
    print(f"  Insert sections (w:ins): {ins_count}")
    print(f"  File size: {os.path.getsize(DST)} bytes")
    print(f"  Output: {DST}")
    print(f"{'='*50}")
    print(f"\nOpen in Word and enable Track Changes / Review Mode to see the redlines!")
PYEOF
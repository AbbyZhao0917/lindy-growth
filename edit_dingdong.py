#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import zipfile
from lxml import etree
import os, shutil

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

src = '/home/admin/lindy-growth/temp.docx'
dst = '/home/admin/lindy-growth/\u53ee\u549a\u4e70\u83dc\u6848\u4f8b_\u4fee\u6539\u7a3f_\u5e26\u6807\u6ce8.docx'

shutil.copy(src, dst)

with zipfile.ZipFile(src, 'r') as zin:
    old_xml = zin.read('word/document.xml')

root = etree.fromstring(old_xml)

def get_para_text(para):
    texts = []
    for t in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)

def set_run_color(run, color_val):
    rpr = run.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    if rpr is None:
        rpr = etree.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        run.insert(0, rpr)
    color = rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
    if color is None:
        color = etree.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
        rpr.append(color)
    color.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', color_val)

def remove_shd(run):
    for shd in list(run.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')):
        if shd.getparent() is not None:
            shd.getparent().remove(shd)

# ===== STEP 1: Convert yellow highlights to blue text =====
print("Step 1: Converting yellow highlights to blue text...")
for para in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
    for run in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'):
        shd = run.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
        if shd is not None:
            fill = shd.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
            if fill == 'FFFF00':
                set_run_color(run, '0055CC')
                remove_shd(run)
        hl = run.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
        if hl is not None:
            val = hl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if val in ('yellow',):
                set_run_color(run, '0055CC')
                hl.getparent().remove(hl)

# ===== STEP 2: Language polish =====
print("Step 2: Polishing language...")

paras = list(root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'))
text_replacements = {}

for idx, para in enumerate(paras):
    text = get_para_text(para)
    
    if idx == 5:
        text_replacements[idx] = (
            '生鲜零售企业天然面临生鲜产品易腐易损、供应链条长、管理环节多的行业特性。'
            '在当前的消费环境下，用户需求已从\u201c价格敏感\u201d转向\u201c品质与安全敏感\u201d；'
            '与此同时，《食品安全法》《电子商务法》及97号令《企业落实食品安全主体责任监督管理规定》'
            '等政策法规持续趋严，不断强化企业主体责任；'
            '加之行业同质化竞争日益加剧，零售企业正面临严峻的食品安全挑战与运营压力。'
        )
    
    elif idx == 6:
        text_replacements[idx] = (
            '叮咚买菜自创立之初即提出\u201c绝不把不好的菜卖给用户\u201d的核心价值观，'
            '突破传统\u201c事后监管\u201d模式，确立\u201c文化先行、预防为主\u201d的治理理念，'
            '将食品安全文化建设提升至战略高度。'
            '企业将食品安全意识系统性地融入采购、仓储、加工、配送等全链条环节，'
            '围绕\u201c打造全员品质文化，实施全链条\u20187+1\u201d品控体系\u201d的质量方针持续推进，'
            '目标在于构建覆盖\u201c从田间到餐桌\u201d的全员参与、全链路渗透的食品安全文化体系，'
            '推动食安管理从\u201c被动合规\u201d向\u201c主动引领\u201d转型，将文化软实力转化为品质硬标准。'
        )
    
    elif idx == 7:
        text_replacements[idx] = (
            '面对日益复杂的食品安全挑战，叮咚买菜以战略高度推进食品安全文化建设，'
            '构建\u201c金字塔\u201d三层架构——塔尖战略引领、塔身载体创新、塔基生态协同——'
            '系统性地将食品安全意识融入企业运营的每一个环节。'
        )
    
    elif idx == 9:
        text_replacements[idx] = (
            '食品安全文化建设是一项意识渗透工程，既需\u201c自上而下\u201d的意识宣导，'
            '也需\u201c自下而上\u201d的行动推进。为实现食品安全行为的全面一致性，'
            '叮咚买菜通过\u201c金字塔式\u201d建设路径系统性地推广食品安全文化。'
        )
    
    elif idx == 11:
        text_replacements[idx] = (
            '自创立以来，叮咚买菜创始人梁昌霖即提出\u201c绝不把不好的菜卖给用户\u201d的核心价值观。'
            '2025年，企业进一步提出\u201c不做所有人的75分，只做少数人的120分\u201d的战略定位，'
            '主动放弃行业内多数企业竞相追逐的低价策略与低价商品路线，'
            '转而以品质商品服务于对功能、品质、稀缺性等方面产品升级仍有明确需求的消费群体。'
        )
    
    elif idx == 13:
        text_replacements[idx] = (
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
    
    elif idx == 15:
        text_replacements[idx] = (
            '叮咚买菜高级管理层全体亲身践行\u201c绝不把不好的菜卖给用户\u201d的核心价值观。'
            '除授予品控中心绝对的\u201c一票否决\u201d权之外，'
            '企业于每年供应商大会设立\u201c品质食力\u00b7年度品质杰出供应商\u201d奖项，'
            '由品控中心全权提名确认，重点表彰品质表现优异的供应商；'
            '每年夏初，高管团队拍摄夏季食安专项行动宣导片，预热并强调夏季食品安全风险，'
            '与公司全员同步各品类商品组的夏季食安行动计划。'
            '在任何关键季节、关键节点及公司策略调整之际，高级管理层始终秉持对品质的敬畏之心，'
            '在品质之路上坚持\u201c一寸窄，一公里深\u201d的经营方向。'
        )
    
    elif idx == 18:
        text_replacements[idx] = (
            '在战略引领的基础上，叮咚买菜着力打造创新的文化传播载体，'
            '通过多元化的内容运营与智能化工具，使食品安全文化触达每一位员工。'
        )
    
    elif idx == 22:
        text_replacements[idx] = (
            '在频道内容方面，设置多重差异化内容方向，鼓励全员投稿，'
            '积极传播食品安全知识与行动成果。'
        )
    
    elif idx == 28:
        text_replacements[idx] = (
            '在栏目功能方面，设置三大模块，涵盖员工食安知识库、内部建议反馈及全员排查反馈入口。'
        )
    
    elif idx == 34:
        text_replacements[idx] = (
            '文化建设的最终目标是实现全员参与、全链路渗透。'
            '叮咚买菜通过丰富的文化活动与技能竞赛，将食品安全文化从企业内部延伸至供应商生态。'
        )
    
    elif idx == 36:
        text_replacements[idx] = (
            '构建全员参与、全链路渗透的食品安全文化体系，'
            '需要借助多元化的形式与方法，提升员工的积极性与参与感，强化文化浸润效果。'
        )
    
    elif idx == 38:
        text_replacements[idx] = (
            '通过组织丰富多样的食品安全文化活动，以高级管理层为示范，'
            '引导全员参与，并同步推进相应的食品安全行动。'
        )
    
    elif idx == 45:
        text_replacements[idx] = (
            '以竞赛形式营造争优争先的团队氛围，为企业长远发展储备专业人才。'
            '技能大赛不局限于品控部门发起，更鼓励业务部门主动组织，'
            '以此促进人员专业技能的提升，激发员工潜能与质量意识。'
        )
    
    elif idx == 51:
        text_replacements[idx] = (
            '经过多年的持续建设与深耕，叮咚买菜的食品安全文化体系已取得显著成效，主要成果如下。'
        )
    
    elif idx == 53:
        text_replacements[idx] = (
            '食品安全是叮咚买菜的核心竞争力之一。'
            '通过系统性、全链路的食品安全文化建设，叮咚买菜不仅在内部建立了高标准的质量管控体系，'
            '更通过创新实践与积极影响，构建了良好的食品安全文化氛围。'
            '以下是叮咚买菜食品安全文化的主要成果与价值输出路径：'
        )

# Apply replacements
for idx, new_text in text_replacements.items():
    para = paras[idx]
    # Remove existing runs
    for r in list(para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')):
        para.remove(r)
    
    # Create new run
    new_run = etree.SubElement(para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rpr = etree.SubElement(new_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    
    color_elem = etree.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
    color_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '333333')
    
    rfonts = etree.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    rfonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', u'\u7b49\u7ebf')
    rfonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', u'\u7b49\u7ebf')
    
    sz = etree.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '24')
    
    new_text_elem = etree.SubElement(new_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    new_text_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space', 'preserve')
    new_text_elem.text = new_text

# Write back to docx
new_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

tmp = dst + '.tmp'
with zipfile.ZipFile(src, 'r') as zin:
    with zipfile.ZipFile(tmp, 'w') as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml':
                data = new_xml
            zout.writestr(item, data)

os.replace(tmp, dst)

# Verify
with zipfile.ZipFile(dst, 'r') as z:
    new_xml_check = z.read('word/document.xml')
    check_root = etree.fromstring(new_xml_check)
    # Count remaining yellow highlights
    yellow_count = 0
    for run in check_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'):
        shd = run.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
        if shd is not None and shd.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill') == 'FFFF00':
            yellow_count += 1
    print(f"Remaining yellow highlights: {yellow_count}")
    
    # Check blue text
    blue_count = 0
    for run in check_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'):
        rpr = run.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        if rpr is not None:
            color = rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
            if color is not None and color.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') == '0055CC':
                blue_count += 1
    print(f"Blue colored runs: {blue_count}")
    
    # Print first few paras to verify
    for i, p in enumerate(list(check_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'))[:12]):
        t = get_para_text(p)
        if t.strip():
            print(f"  P{i}: {t[:120]}...")

print(f"\nDone! Output saved to: {dst}")

"""Render the source-checked literature records as a compact, searchable PDF.

Requires PyMuPDF, fontTools, and the Noto Sans CJK system font. This script renders the
stored research records; it does not perform new searches or verify literature.
"""
from pathlib import Path
from collections import Counter
import json,re,math
import fitz
from fontTools.ttLib import TTCollection
from fontTools import subset

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'文献资料'
RECORDS=json.loads((OUT/'文献条目.json').read_text())
MEDIA=json.loads((OUT/'公众号线索.json').read_text())
def prepare_font(source,name):
    # Extract the SC face and subset before PDF embedding. Passing a TTC to the
    # PDF's post-hoc subsetter can leave invalid CFF charstrings in some versions.
    fonts=TTCollection(source)
    font=fonts.fonts[2]
    options=subset.Options()
    # Noto's CID-keyed CFF fonts require stable glyph IDs for PDF Identity-H.
    options.retain_gids=True
    options.layout_features=['*']
    sub=subset.Subsetter(options=options)
    all_text=Path(__file__).read_text()+json.dumps(RECORDS,ensure_ascii=False)+json.dumps(MEDIA,ensure_ascii=False)
    sub.populate(text=all_text)
    sub.subset(font)
    target=Path('/tmp')/(name+'.otf')
    font.save(target)
    return str(target)

REGULAR=prepare_font('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc','literature_review_regular')
BOLD=prepare_font('/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc','literature_review_bold')
FR=fitz.Font(fontfile=REGULAR)
FB=fitz.Font(fontfile=BOLD)
W,H=fitz.paper_size('a4-l')
M=25
BOTTOM=H-32
INK=(.10,.18,.18)
GREEN=(.13,.34,.29)
MUTED=(.34,.40,.40)
GRID=(.75,.80,.78)
PALE=(.94,.96,.95)
WHITE=(1,1,1)

doc=fitz.open()
page=None
y=0
LOG=[]

def wrap(s,width,size,bold=False):
    f=FB if bold else FR
    result=[]
    for para in str(s).split('\n'):
        if not para:
            result.append('');continue
        # Keep ordinary Latin words intact, while allowing CJK wrapping.
        tokens=re.findall(r'[A-Za-z0-9]+(?:[’\'._:/%+-][A-Za-z0-9]+)*|[ \t]+|[^\x00-\x7F]|.',para)
        line=''
        for token in tokens:
            if f.text_length(line+token,fontsize=size)<=width:
                line+=token;continue
            if line.strip():result.append(line.rstrip())
            line=token.lstrip()
            while line and f.text_length(line,fontsize=size)>width:
                cut=1
                while cut<len(line) and f.text_length(line[:cut+1],fontsize=size)<=width:cut+=1
                result.append(line[:cut]);line=line[cut:]
        if line.strip():result.append(line.rstrip())
    return result or ['']

def put(x,top,lines,size=9.1,bold=False,color=INK,leading=None):
    leading=leading or size*1.35
    for i,line in enumerate(lines):
        baseline=top+size*1.05+i*leading
        page.insert_text((x,baseline),line,fontname='Fbold' if bold else 'Fregular',fontsize=size,color=color)
    return len(lines)*leading

def new_page():
    global page,y
    page=doc.new_page(width=W,height=H)
    page.insert_font(fontname='Fregular',fontfile=REGULAR)
    page.insert_font(fontname='Fbold',fontfile=BOLD)
    put(M,13,['河西走廊梭梭林退化｜国内外文献研究现状'],8.2,color=MUTED)
    page.draw_line((M,28),(W-M,28),color=GREEN,width=.7)
    y=37

def ensure(height):
    if y+height>BOTTOM:new_page()

def para(s,size=9.4,color=INK,bold=False,gap=5):
    global y
    lines=wrap(s,W-2*M,size,bold)
    height=len(lines)*size*1.4
    ensure(height+gap)
    y+=put(M,y,lines,size,bold,color,size*1.4)+gap

def heading(s,sub=None):
    global y
    ensure(65)
    y+=5
    y+=put(M,y,wrap(s,W-2*M,13.2,True),13.2,True,GREEN,18)+5
    if sub:para(sub,8.7,MUTED,gap=6)

COLS=[135,98,100,121,201,W-2*M-655]
HEAD=['文献 / 地区与对象','研究问题与思路','评价内容与指标','数据与方法','主要结果（按文献转述）','证据边界 / 读取程度']

def table_header(widths,headers):
    global y
    height=25
    page.draw_rect(fitz.Rect(M,y,W-M,y+height),color=GREEN,fill=GREEN)
    x=M
    for width,s in zip(widths,headers):
        put(x+5,y+6,wrap(s,width-10,9.0,True),9.0,True,WHITE,11)
        x+=width
    y+=height

def table_row(cells,widths,index,headers,size=9.1):
    global y
    lines=[wrap(s,w-11,size) for s,w in zip(cells,widths)]
    leading=12.1
    height=max(map(len,lines))*leading+13
    assert height<440,(cells[0],height)
    if y+height>BOTTOM:
        new_page();table_header(widths,headers)
    if index%2==0:page.draw_rect(fitz.Rect(M,y,W-M,y+height),color=None,fill=PALE)
    x=M
    for j,(width,ls) in enumerate(zip(widths,lines)):
        page.draw_rect(fitz.Rect(x,y,x+width,y+height),color=GRID,width=.3)
        put(x+5.5,y+6,ls,size,bold=False,leading=leading)
        x+=width
    LOG.append({'page':page.number+1,'kind':'row','id':cells[0].split('\n')[0],'top':y,'bottom':y+height})
    y+=height

def author_label(r):
    if r['id'] in ['A04','A11']:return r['authors'][0]+'等'
    a=r['authors'][0]
    # Take family name from the already checked bibliographic record.
    if r['id']=='B04':return 'Yang等'
    if r['id']=='C01':return 'Ahn等'
    if r['id']=='C02':return 'Tsedensodnom等'
    return a.split()[-1]+'等'

new_page()
para('河西走廊梭梭林退化：国内外文献研究现状表',18.5,GREEN,True,gap=6)
para('定向检索整理｜检索截至2026年10月15日｜24篇学术文献＋4条微信公众号索引线索',10.2,MUTED,gap=5)
para('阅读标记：全文核心段＝已核对原文方法与结果相关段；期刊摘要＝已读期刊官网摘要；数据库摘要＝已读Crossref注册摘要或OpenAlex索引摘要。公众号仅列已核实的索引信息，未读到的原文不作为科学结论。',8.8,MUTED,gap=4)
para('分类按研究地区和对象，英文发表的中国研究仍列国内。“证据边界”是本次评述，不全部等同作者自述不足。本文为定向文献整理，并非穷尽式系统综述；完整题名、题录和可点击来源见文后。',8.8,MUTED,gap=8)

groups=list(dict.fromkeys(r['group'] for r in RECORDS))
for group in groups:
    rs=[r for r in RECORDS if r['group']==group]
    # Keep heading, column header and at least a meaningful start together.
    ensure(160)
    heading(group+f'（{len(rs)}篇）')
    table_header(COLS,HEAD)
    for i,r in enumerate(rs):
        first=f"[{r['id']}] {author_label(r)}，{r['year']}\n{r['short_title']}\n{r['where']}"
        last=r['limit']+'\n读取：'+r['read_level']+'。'
        table_row([first,r['question'],r['metrics'],r['method'],r['result'],last],COLS,i,HEAD)
    y+=7

heading('五、从这些文献可以归纳出的研究进展')
summary=[
('河西走廊已研究到哪些层次？','已覆盖生态因子与退化状态、水分与生理响应、空间结构和自然更新、林分耗水、碳储量及土壤动物等。民勤和临泽存在直接实证资料，并非仅有造林成活调查。','A11、A01–A09（A03不计精确地名）'),
('不同论文为什么会得出不同的“老林表现”？','有的研究显示老龄林表土改善或部分光合能力回升，有的显示更新、地上生物量或深层水分下降。研究对象、选样标准和指标不同，不能把土壤改善等同于植株健康。','A01、A02、A04、B01、B02'),
('土壤含水率能否形成一个统一阈值？','目前所列研究的数值来自不同土壤、土层和水势定义：0.824%/1.30%、约1%–2%、经济阈值约2%与生存阈值约1%含义不相同，不能简单平均或直接通用。','A11、A04、B04'),
('衰退原因已经解释到什么程度？','已有水分来源、土壤干化、空间竞争、风沙伤害和微生物关联等证据；其中包含观察推断，也有风洞、疏伐等处理研究。不同证据的因果强度需分别判断。','A03、B01、B03、B05、D02、D04'),
('未来风险与长期恢复效果证据如何？','本次梭梭文献以年龄序列、过程监测和早期建植试验为主；未据此确认覆盖四地的同株长期衰退预测系统。柠条2年处理试验等资料提供恢复效果证据，但并非梭梭通用处方。','C01、D01、D02；本次纳入范围'),
]
sumwidth=[145,W-2*M-270,125]
sumhead=['文献问题','已有认识及范围','对应文献']
table_header(sumwidth,sumhead)
for i,cells in enumerate(summary):table_row(cells,sumwidth,i,sumhead,size=9.2)
y+=5
para('地区覆盖说明：本表河西走廊直接证据主要来自民勤、石羊河下游与张掖临泽。A03摘要未明确具体地名，未用于地区覆盖统计；没有将未核到的酒泉/敦煌论文写作“不存在研究”。',8.8,MUTED,gap=7)

ensure(200)
heading('六、微信公众号与传播资料线索（不混入论文结果）',
        '以下题名、账号、日期及片段已在搜狗微信索引读取；微信原文受访问限制/验证码影响，全文与原始论文追溯尚未完成。日期为索引显示的日期；链接为检索入口，不冒充原文永久链接。')
mediawidth=[183,109,250,W-2*M-542]
mediahead=['条目与索引日期','公众号','可参考内容与使用边界','来源入口 / 核验程度']
table_header(mediawidth,mediahead)
for i,r in enumerate(MEDIA):
    top=y
    table_row([f"[{r['id']}] {r['title']}\n{r['date']}",r['account'],r['use'],'搜狗微信检索入口（点击本格）\n仅索引已核，原文未核；不据片段引用具体效应或阈值。'],mediawidth,i,mediahead,size=9.2)
    # The row may have moved to a new page; use its actual geometry.
    top=LOG[-1]['top'];xx=M+sum(mediawidth[:3])
    page.insert_link({'kind':fitz.LINK_URI,'from':fitz.Rect(xx,top,W-M,y),'uri':r['url']})
y+=8

heading('七、完整文献题录与原文入口',
        '编号与主表对应。英文题名保留原文，主表短标题为中文概括；日期采用卷期年，并注明与在线年不一致的条目。除A11未核到DOI外，其余23篇DOI均已读取注册题录。')

for r in RECORDS:
    authors=r['authors'][:3]
    auth='; '.join(authors)+('; et al.' if len(r['authors'])>3 else '')
    vol=r['volume']+(f"({r['issue']})" if r['issue'] else '')
    cite=f"[{r['id']}] {auth}. {r['year']}. {r['title']}. {r['journal']}"
    if vol:cite+=', '+vol
    if r['pages']:cite+=': '+r['pages']
    cite+='.'
    if r['id']=='C02':cite+=' [会议论文，ESTIC 2021]'
    if r['id']=='A03':cite+=' [在线2016，卷期2017]'
    if r['id']=='A06':cite+=' [在线2022，卷期2023]'
    if r['id']=='A07':cite+=' [在线2023，卷期2024]'
    if r['id']=='D02':cite+=' [在线2022，卷期2023]'
    if r['id']=='B07':cite+=' [2024年发表，卷标2023]'
    cite_lines=wrap(cite,W-2*M,9.1)
    link=('DOI: https://doi.org/'+r['doi']) if r['doi'] else ('期刊链接: '+r['primary_url'])
    link_lines=wrap(link,W-2*M,8.7)
    height=len(cite_lines)*12.0+len(link_lines)*11.5+10
    ensure(height)
    y+=put(M,y,cite_lines,9.1,leading=12.0)
    linktop=y
    y+=put(M,y,link_lines,8.7,color=GREEN,leading=11.5)
    page.insert_link({'kind':fitz.LINK_URI,'from':fitz.Rect(M,linktop,W-M,y),'uri':('https://doi.org/'+r['doi']) if r['doi'] else r['primary_url']})
    y+=9

heading('检索与核验简记')
para('检索使用Crossref、OpenAlex及可访问的期刊官网，组合Haloxylon ammodendron/梭梭、Hexi/河西走廊、Minqin/民勤、Zhangye/张掖、Jiuquan/酒泉、Dunhuang/敦煌与decline、mortality、survival、soil water、regeneration、衰退、退化等词，并追踪相关题录。排除未核题录的泛化文章、与主题无关的检索结果及未核实的预印本。',8.8,MUTED)
para('全文读取并不覆盖所有论文；读取程度在每行单列。公众号索引只能证明该线索被检索到，不能替代原文核验。本文未复用用户先前撤回的论文示例数值，未编造当地保存率或调查数据。',8.8,MUTED)

for i,p in enumerate(doc):
    p.draw_line((M,H-25),(W-M,H-25),color=GRID,width=.5)
    p.insert_text((M,H-12),'文献研究现状表｜定向检索整理｜2026-10-15',fontname='Fregular',fontsize=7.4,color=MUTED)
    label=f'{i+1} / {len(doc)}'
    length=FR.text_length(label,fontsize=8)
    p.insert_text((W-M-length,H-12),label,fontname='Fregular',fontsize=8,color=MUTED)

doc.set_metadata({'title':'河西走廊梭梭林退化：国内外文献研究现状表','subject':'24篇学术文献的逐篇整理与4条微信公众号索引线索','author':'','keywords':'梭梭;河西走廊;文献;退化;保存;更新'})
pdf=OUT/'河西走廊梭梭林退化_国内外文献研究现状.pdf'
doc.save(pdf,garbage=4,deflate=True)
(OUT/'PDF排版记录.json').write_text(json.dumps({'pages':len(doc),'academic_records':len(RECORDS),'media_leads':len(MEDIA),'rows':LOG},ensure_ascii=False,indent=2))
print(json.dumps({'pdf':str(pdf),'pages':len(doc),'academic_records':len(RECORDS),'media_leads':len(MEDIA),'reading_levels':Counter(r['read_level'] for r in RECORDS)},ensure_ascii=False))

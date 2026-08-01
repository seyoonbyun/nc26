import json

with open(r'C:\Users\Gram\desktop\typebot-export-dna-cjrblax.json','r',encoding='utf-8') as f:
    tpl = json.load(f)

def txt(bid, text):
    return {'id':bid,'type':'text','content':{'richText':[{'type':'p','children':[{'text':text}]}],'plainText':text}}

def rate(bid, vid, btn):
    return {'id':bid,'type':'choice input','items':[{'id':f'{bid}_{i}','content':str(i),'value':str(i)} for i in range(1,6)],'options':{'variableId':vid,'isMultipleChoice':False,'buttonLabel':btn}}

def ch(bid, vid, btn, pairs):
    return {'id':bid,'type':'choice input','items':[{'id':f'{bid}_{i}','content':p[0],'value':p[1]} for i,p in enumerate(pairs)],'options':{'variableId':vid,'isMultipleChoice':False,'buttonLabel':btn}}

def inp(bid, vid, ph, btn, long=False):
    o = {'labels':{'placeholder':ph,'button':btn},'variableId':vid}
    if long: o['isLong'] = True
    return {'id':bid,'type':'text input','options':o}

def nps_block(bid, vid, btn):
    return {'id':bid,'type':'choice input','items':[{'id':f'{bid}_{i}','content':str(i),'value':str(i)} for i in range(11)],'options':{'variableId':vid,'isMultipleChoice':False,'buttonLabel':btn}}

rv = ['v06','v07','v08','v09','v10','v11','v12','v13','v14','v15','v16','v17','v18','v20','v21','v22','v23','v24','v25']

KR=['전반적 만족도','기대 충족도','[Day1] 존윤 오프닝','멤버 성공 스토리','글로벌 스토리','김하원 키노트','Fujimoto 스피치','파티 & 어워드','[Day2] Judee Quiazon','Mac Srinivasan','지역 발표','Norm Dominguez','시상식','[장소] 시설','등록/체크인','동선','식음료','부스','네트워킹']
EN=['Overall satisfaction','Expectations','[Day1] John Yoon','Member Success Stories','Global Story','Kim Hawon Keynote','Fujimoto Speech','Party & Awards','[Day2] Judee Quiazon','Mac Srinivasan','Regional Presentations','Norm Dominguez','Award Ceremony','[Venue] Facilities','Registration','Event Flow','Food & Beverage','Booth','Networking']
JP=['全体満足度','期待充足度','[Day1] John Yoon','サクセスストーリー','グローバルストーリー','Kim Hawon','Fujimoto','パーティー','[Day2] Judee Quiazon','Mac Srinivasan','地域プレゼン','Norm Dominguez','授賞式','[会場] 施設','受付','動線','飲食','ブース','ネットワーキング']
CN=['整体满意度','期望满足度','[Day1] John Yoon','成功故事','全球故事','Kim Hawon','Fujimoto','派对','[Day2] Judee Quiazon','Mac Srinivasan','地区展示','Norm Dominguez','颁奖典礼','[场地] 设施','注册','动线','餐饮','展位','交流']

best_labels = {'k':['존윤 오프닝','성공 스토리','글로벌 스토리','김하원 키노트','Fujimoto','파티&어워드','Judee Quiazon','Mac Srinivasan','지역 발표','Norm Dominguez'],
               'e':['John Yoon','Success Stories','Global Story','Kim Hawon','Fujimoto','Party&Awards','Judee Quiazon','Mac Srinivasan','Regional','Norm Dominguez'],
               'j':['John Yoon','サクセス','グローバル','Kim Hawon','Fujimoto','パーティー','Judee Quiazon','Mac Srinivasan','地域プレゼン','Norm Dominguez'],
               'c':['John Yoon','成功故事','全球故事','Kim Hawon','Fujimoto','派对','Judee Quiazon','Mac Srinivasan','地区展示','Norm Dominguez']}

def build(p, labels, btn, intro, nph, eph, cph, att, day, npstxt, again, bst_q, q1, q2, q3, phf, btns):
    b = [txt(f'{p}01', intro), inp(f'{p}02','v01',nph,btn), inp(f'{p}03','v02',eph,btn), inp(f'{p}04','v03',cph,btn)]
    b.append(ch(f'{p}05','v04',btn,att))
    b.append(ch(f'{p}06','v05',btn,day))
    b.append(txt(f'{p}07', labels[0]))
    b.append(rate(f'{p}r0', rv[0], btn))
    for i in range(1, len(rv)):
        b.append(txt(f'{p}r{i}t', labels[i]))
        b.append(rate(f'{p}r{i}', rv[i], btn))
    b.append(txt(f'{p}bs_t', bst_q))
    bp = [(x,x) for x in best_labels[p]]
    b.append(ch(f'{p}bs','v19',btn,bp))
    b.append(txt(f'{p}nps_t', npstxt))
    b.append(nps_block(f'{p}nps','v26',btn))
    b.append(ch(f'{p}att','v27',btn,again))
    b.append(txt(f'{p}ft1', q1)); b.append(inp(f'{p}fi1','v28',phf,btn,True))
    b.append(txt(f'{p}ft2', q2)); b.append(inp(f'{p}fi2','v29',phf,btn,True))
    b.append(txt(f'{p}ft3', q3))
    last = inp(f'{p}fi3','v30',phf,btns,True)
    last['outgoingEdgeId'] = f'e_{p}_sub'
    b.append(last)
    return b

krB = build('k',KR,'다음','안녕하세요! NC26 참석 감사합니다. 설문 참여 부탁드립니다 (약3분)','홍길동','email@example.com','예: 강남 플러스',[('멤버','멤버'),('비지터','비지터'),('게스트','게스트'),('리더십','리더십')],[('Day1만','Day1'),('Day2만','Day2'),('양일','Both')],'[추천] 동료에게 추천? (0=전혀, 10=매우)',[('반드시','반드시'),('아마도','아마도'),('모르겠음','모르겠음'),('아니요','아니요')],'가장 인상 깊었던 세션?','가장 좋았던 점?','개선할 점?','추가 의견?','자유롭게 작성','제출')
enB = build('e',EN,'Next','Hello! Thank you for attending NC26. (~3 min)','Your Name','email@example.com','e.g. Gangnam Plus',[('Member','Member'),('Visitor','Visitor'),('Guest','Guest'),('Leadership','Leadership')],[('Day1 only','Day1'),('Day2 only','Day2'),('Both','Both')],'[NPS] Recommend? (0=Not at all, 10=Strongly)',[('Definitely','Definitely'),('Maybe','Maybe'),('Not Sure','Not Sure'),('No','No')],'Most impressive session?','Best part?','Suggestions?','Other comments?','Write freely','Submit')
jpB = build('j',JP,'次へ','こんにちは！NC26ご参加ありがとうございます（約3分）','お名前','email@example.com','例: Tokyo Central',[('メンバー','メンバー'),('ビジター','ビジター'),('ゲスト','ゲスト'),('リーダーシップ','リーダーシップ')],[('Day1のみ','Day1'),('Day2のみ','Day2'),('両日','Both')],'[推薦] 同僚に勧める?(0=全く,10=強く)',[('必ず参加','必ず参加'),('おそらく','おそらく'),('わからない','わからない'),('参加しない','参加しない')],'最も印象的なセッション?','最も良かった点?','改善点?','その他?','ご自由に','送信')
cnB = build('c',CN,'下一步','您好！感谢参加NC26（约3分钟）','您的姓名','email@example.com','例: HK Central',[('会员','会员'),('访客','访客'),('嘉宾','嘉宾'),('领导层','领导层')],[('仅Day1','Day1'),('仅Day2','Day2'),('两天','Both')],'[推荐] 推荐?(0=完全不,10=强烈)',[('一定参加','一定参加'),('可能会','可能会'),('不确定','不确定'),('不会','不会')],'最深印象的场次?','最喜欢的?','改进建议?','其他?','请自由填写','提交')

tpl['id']='nc26survey2026'; tpl['name']='NC26 Satisfaction Survey'; tpl['publicId']=None; tpl['customDomain']=None
tpl['events']=[{'id':'ev1','type':'start','graphCoordinates':{'x':0,'y':0},'outgoingEdgeId':'e_start'}]
tpl['groups']=[
  {'id':'g_lang','title':'Language','graphCoordinates':{'x':250,'y':0},'blocks':[
    txt('lt','Select language / 언어 선택 / 言語選択 / 选择语言'),
    ch('lc','v_lang','Select',[('한국어','한국어'),('English','English'),('日本語','日本語'),('中文','中文')]),
    {'id':'lcond','type':'Condition','items':[
      {'id':'ci_kr','outgoingEdgeId':'e_kr','content':{'logicalOperator':'AND','comparisons':[{'id':'cp_kr','variableId':'v_lang','comparisonOperator':'Equal to','value':'한국어'}]}},
      {'id':'ci_en','outgoingEdgeId':'e_en','content':{'logicalOperator':'AND','comparisons':[{'id':'cp_en','variableId':'v_lang','comparisonOperator':'Equal to','value':'English'}]}},
      {'id':'ci_jp','outgoingEdgeId':'e_jp','content':{'logicalOperator':'AND','comparisons':[{'id':'cp_jp','variableId':'v_lang','comparisonOperator':'Equal to','value':'日本語'}]}}]},
    {**txt('ldef',''),'outgoingEdgeId':'e_cn'}]},
  {'id':'gk','title':'KR','graphCoordinates':{'x':250,'y':200},'blocks':krB},
  {'id':'ge','title':'EN','graphCoordinates':{'x':700,'y':200},'blocks':enB},
  {'id':'gj','title':'JP','graphCoordinates':{'x':1150,'y':200},'blocks':jpB},
  {'id':'gc','title':'CN','graphCoordinates':{'x':1600,'y':200},'blocks':cnB},
  {'id':'g_thx','title':'Thanks','graphCoordinates':{'x':900,'y':1600},'blocks':[txt('thx','Thank you! 감사합니다!\nwww.nc26-bnikorea.com')]}]
tpl['edges']=[
  {'id':'e_start','from':{'eventId':'ev1'},'to':{'groupId':'g_lang'}},
  {'id':'e_kr','from':{'blockId':'lcond'},'to':{'groupId':'gk'}},
  {'id':'e_en','from':{'blockId':'lcond'},'to':{'groupId':'ge'}},
  {'id':'e_jp','from':{'blockId':'lcond'},'to':{'groupId':'gj'}},
  {'id':'e_cn','from':{'blockId':'ldef'},'to':{'groupId':'gc'}},
  {'id':'e_k_sub','from':{'blockId':'kfi3'},'to':{'groupId':'g_thx'}},
  {'id':'e_e_sub','from':{'blockId':'efi3'},'to':{'groupId':'g_thx'}},
  {'id':'e_j_sub','from':{'blockId':'jfi3'},'to':{'groupId':'g_thx'}},
  {'id':'e_c_sub','from':{'blockId':'cfi3'},'to':{'groupId':'g_thx'}}]
tpl['variables']=[
  {'id':'v_lang','name':'lang','isSessionVariable':False,'value':None},
  {'id':'v01','name':'name','isSessionVariable':False,'value':None},
  {'id':'v02','name':'email','isSessionVariable':False,'value':None},
  {'id':'v03','name':'chapter','isSessionVariable':False,'value':None},
  {'id':'v04','name':'att_type','isSessionVariable':False,'value':None},
  {'id':'v05','name':'days','isSessionVariable':False,'value':None},
  {'id':'v06','name':'overall','isSessionVariable':False,'value':None},
  {'id':'v07','name':'expect','isSessionVariable':False,'value':None},
  {'id':'v08','name':'d1_opening','isSessionVariable':False,'value':None},
  {'id':'v09','name':'d1_success','isSessionVariable':False,'value':None},
  {'id':'v10','name':'d1_global','isSessionVariable':False,'value':None},
  {'id':'v11','name':'d1_keynote','isSessionVariable':False,'value':None},
  {'id':'v12','name':'d1_fujimoto','isSessionVariable':False,'value':None},
  {'id':'v13','name':'d1_party','isSessionVariable':False,'value':None},
  {'id':'v14','name':'d2_judee','isSessionVariable':False,'value':None},
  {'id':'v15','name':'d2_mac','isSessionVariable':False,'value':None},
  {'id':'v16','name':'d2_regional','isSessionVariable':False,'value':None},
  {'id':'v17','name':'d2_norm','isSessionVariable':False,'value':None},
  {'id':'v18','name':'d2_award','isSessionVariable':False,'value':None},
  {'id':'v19','name':'best_session','isSessionVariable':False,'value':None},
  {'id':'v20','name':'venue','isSessionVariable':False,'value':None},
  {'id':'v21','name':'registration','isSessionVariable':False,'value':None},
  {'id':'v22','name':'flow','isSessionVariable':False,'value':None},
  {'id':'v23','name':'food','isSessionVariable':False,'value':None},
  {'id':'v24','name':'booth','isSessionVariable':False,'value':None},
  {'id':'v25','name':'networking','isSessionVariable':False,'value':None},
  {'id':'v26','name':'nps','isSessionVariable':False,'value':None},
  {'id':'v27','name':'attend_again','isSessionVariable':False,'value':None},
  {'id':'v28','name':'best_part','isSessionVariable':False,'value':None},
  {'id':'v29','name':'suggestions','isSessionVariable':False,'value':None},
  {'id':'v30','name':'comments','isSessionVariable':False,'value':None}]

with open(r'C:\Users\Gram\desktop\nc26\typebot-nc26-survey.json','w',encoding='utf-8') as f:
    json.dump(tpl, f, ensure_ascii=False, separators=(',',':'))

import os
s = os.path.getsize(r'C:\Users\Gram\desktop\nc26\typebot-nc26-survey.json')
nb = sum(len(g['blocks']) for g in tpl['groups'])
print(f'Done. Size={s}B Groups={len(tpl["groups"])} Blocks={nb} Edges={len(tpl["edges"])} Vars={len(tpl["variables"])}')

#!/bin/bash
# 2-pass build: pass1 → page numbers → pass2 with TOC
set -e
B=/tmp/claude-0/-home-user-claude/ec329815-8553-5a66-b013-4e0ab91a067d/scratchpad/build
O=/tmp/claude-0/-home-user-claude/ec329815-8553-5a66-b013-4e0ab91a067d/scratchpad/out
SOFFICE=/root/.claude/skills/synced/32fb2ddc-a363-472f-9f70-6ddfa593095a_d15c34d7-6182-490b-a97b-4c2ad6008ce8/docx/scripts/office/soffice.py
NAME=$(grep -o "2027_BizPlan_PreSurvey_MfgAX_v[0-9.]*\.docx" $B/build.js | head -1)
cd $B
NO_H1_BREAK=1 node build.js >/dev/null
cd $O && python3 $SOFFICE --headless --convert-to pdf $NAME >/dev/null 2>&1
PDF=${NAME%.docx}.pdf
python3 -I - "$B" "$O/$PDF" <<'EOF'
import subprocess, json, re, sys
B, pdf = sys.argv[1], sys.argv[2]
n=int(re.search(r'Pages:\s+(\d+)',subprocess.run(['pdfinfo',pdf],capture_output=True,text=True).stdout).group(1))
pages=[subprocess.run(['pdftotext','-f',str(i),'-l',str(i),'-layout',pdf,'-'],capture_output=True,text=True).stdout for i in range(1,n+1)]
heads=[]
for ch in ['ch1','ch2','ch3','ch4']:
    t=open(f'{B}/{ch}.js',encoding='utf8').read()
    for m in re.finditer(r"\{ h: ([12]), t: '([^']+)' \}",t): heads.append((int(m.group(1)),m.group(2)))
ga='가나다라마바사아자차카타파하'; n1=n2=0; res={}; start=3; missing=[]
norm=lambda x: re.sub(r'\s+','',x)
for lv,t in heads:
    if lv==1: n1+=1; n2=0; label=f'{n1}. {t}'
    else: n2+=1; label=f'{ga[n2-1]}. {t}'
    found=None
    for i in range(start-1,n):
        if norm(label) in norm(pages[i]): found=i+1; break
    if found is None: missing.append(label)
    else: res[f'{lv}:{t}']=found; start=found
json.dump(res,open(f'{B}/toc_pages.json','w',encoding='utf8'),ensure_ascii=False)
print('pass1 pages', n, 'missing', missing)
EOF
cd $B && TOC_PAGES="$(cat toc_pages.json)" NO_H1_BREAK=1 node build.js >/dev/null
cd $O && python3 $SOFFICE --headless --convert-to pdf $NAME >/dev/null 2>&1
echo "final: $(pdfinfo $PDF | grep Pages)"

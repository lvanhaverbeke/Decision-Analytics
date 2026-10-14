from pathlib import Path
import re,html
P=Path(__file__).resolve().parent
src=(P/'executive_report.md').read_text().replace('€33,697','€33,696')
(P/'executive_report.md').write_text(src)
def inline(s):
 s=html.escape(s)
 s=re.sub(r'\[([^\]]+)\]\((https?://[^)]+)\)',r'<a href="\2">\1</a>',s)
 s=re.sub(r'\*\*([^*]+)\*\*',r'<strong>\1</strong>',s)
 return re.sub(r'`([^`]+)`',r'<code>\1</code>',s)
out=[];table=False;ul=False
for line in src.splitlines():
 if not line:
  if table:out.append('</tbody></table>');table=False
  if ul:out.append('</ul>');ul=False
  continue
 if line.startswith('|'):
  cells=[x.strip() for x in line.strip('|').split('|')]
  if all(re.match(r'^:?-+:?$',x) for x in cells):continue
  if not table:out.append('<table><tbody>');table=True
  out.append('<tr>'+''.join('<td>'+inline(x)+'</td>' for x in cells)+'</tr>');continue
 if line.startswith('- '):
  if not ul:out.append('<ul>');ul=True
  out.append('<li>'+inline(line[2:])+'</li>');continue
 if line.startswith('# '):out.append('<h1>'+inline(line[2:])+'</h1>')
 elif line.startswith('## '):out.append('<h2>'+inline(line[3:])+'</h2>')
 else:out.append('<p>'+inline(line)+'</p>')
(P/'executive_report.html').write_text('<!doctype html><html><head><meta charset="utf-8"><title>Student Housing Executive Report</title><style>body{font:15px/1.55 Arial;max-width:950px;margin:40px auto;color:#243746}h1,h2{color:#183b56}table{border-collapse:collapse;width:100%}td{border-bottom:1px solid #ddd;padding:7px}tr:first-child{background:#183b56;color:white;font-weight:bold}td:last-child{text-align:right}a{color:#16638d}@media print{body{margin:10mm;font-size:10pt}h2{break-after:avoid}tr{break-inside:avoid}}</style></head><body>'+''.join(out)+'</body></html>')
print('Formatted executive report saved. SAVE_NOW:',(P/'SAVE_NOW.txt').exists())

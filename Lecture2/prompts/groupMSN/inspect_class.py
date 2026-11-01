import struct as s,json
p='Chapter 1 - New road analysis model.xls';b=open(p,'rb').read();ss=1<<s.unpack_from('<H',b,30)[0]
def sec(i):return b[(i+1)*ss:(i+2)*ss]
fatids=[i for i in s.unpack_from('<109I',b,76) if i<0xfffffffa];fat=[]
for i in fatids:fat+=list(s.unpack('<%dI'%(ss//4),sec(i)))
def chain(i):
 a=b'';seen=set()
 while i<0xfffffffa and i not in seen:seen.add(i);a+=sec(i);i=fat[i]
 return a
d=chain(s.unpack_from('<I',b,48)[0]);streams={}
for o in range(0,len(d),128):
 n=s.unpack_from('<H',d,o+64)[0]
 if n>2:
  name=d[o:o+n-2].decode('utf16');start=s.unpack_from('<I',d,o+116)[0];size=s.unpack_from('<Q',d,o+120)[0];streams[name]=(start,size)
print(streams)
a=chain(streams.get('Workbook',streams.get('Book'))[0]);recs=[];o=0
while o+4<=len(a):
 t,l=s.unpack_from('<HH',a,o);r=a[o+4:o+4+l];recs.append((o,t,r));o+=4+l
sst=[]
for o,t,r in recs:
 if t==0xfc:
  n=s.unpack_from('<I',r,4)[0];i=8
  for j in range(n):
   try:
    ln,fl=s.unpack_from('<HB',r,i);i+=3;nr=0;ex=0
    if fl&8:nr=s.unpack_from('<H',r,i)[0];i+=2
    if fl&4:ex=s.unpack_from('<I',r,i)[0];i+=4
    size=ln*(2 if fl&1 else 1);v=r[i:i+size].decode('utf-16le' if fl&1 else 'latin1');i+=size+nr*4+ex;sst.append(v)
   except:break
bounds=[]
for o,t,r in recs:
 if t==0x85:
  off=s.unpack_from('<I',r)[0];ln=r[6];fl=r[7];name=r[8:8+ln*(2 if fl&1 else 1)].decode('utf-16le' if fl&1 else 'latin1');bounds.append((off,name))
print('SHEETS',bounds,'SST',len(sst));current='Globals';out=[]
for o,t,r in recs:
 for off,name in bounds:
  if off==o:current=name;print('\nSHEET',name)
 if t==0xfd:
  row,col,xf,idx=s.unpack('<HHHI',r)
  if idx<len(sst):print(row+1,col+1,sst[idx]);out.append((current,row+1,col+1,sst[idx]))
 elif t==0x203:
  row,col,xf,num=s.unpack('<HHHd',r);print(row+1,col+1,num)
 elif t==6:
  row,col,xf=s.unpack_from('<HHH',r);val=s.unpack_from('<d',r,6)[0];tok=r[22:].hex();print('FORMULA',row+1,col+1,'cached',val,'tokens',tok)
json.dump(out,open('/private/tmp/classread.json','w'))

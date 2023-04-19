import os
from collections import deque
import time
def csv_split(file,batch,target,hasHead=False)->bool:
    filename=os.path.basename(file).split('.')[0]
    row ,count=0,1
    buffer=''
    with open(file) as f:
        if hasHead:
            data=f.read(256)      # 假设标题小于256字节
            lt=data.split('\n')
            head=lt[0]
            buffer+='\n'.join(lt[1:])
        while True:
            if row==0:
                newfile=target+filename+'-'+str(count)+'.csv'
                out=open(newfile,"a+")
                out.write(head+'\n')
            data=f.read(2048)
            if not data:
                break
            buffer+=data
            dq = deque(buffer.split('\n'))
            buffer = ''
            while len(dq) != 1:
                wd = dq.popleft()
                out.write(wd + '\n')
                row += 1
                if row == batch:
                    row = 0
                    count += 1
                    out.close()
                    newfile = target + filename + '-' + str(count) + '.csv'
                    out = open(newfile, "a+")
                    out.write(head + '\n')
            buffer+=dq.popleft()
        if buffer != "":
            out.write(buffer)
            out.close()


csv_split("副本拆封.csv",100000,"D:/pythonProject/MRobotPackage/ExcelMethods/",hasHead=True)

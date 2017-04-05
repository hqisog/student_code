import os
def eachFile(filepathc):
    pathdir=os.listdir(filepathc)
    for alldir in pathdir:
        childdir=os.path.join('%s%s'%(filepathc,alldir))
        print childdir.decode('gbk')

if __name__=='__main__':
    filepathc="/home/hanqing/django_workspace"
    eachFile(filepathc)
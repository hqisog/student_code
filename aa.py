a = 1
b = [2, 3]

def func():
    global a
    a = 2
    print "in func a:", a
    b[0] = 1
    print "in func b:", b

if __name__ == '__main__':
    print "before func a:", a
    print "before func b:", b
    func()
    print "after func a:", a
    print "after func b:", b
    a=a-1
    print 'last  ',a
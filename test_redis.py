# import redis
#
# pool = redis.ConnectionPool(host='127.0.0.1', port=6379)
# r = redis.Redis(connection_pool=pool)
#
# if __name__=='__main__':
#     r.set('name', 'zhangsan')
#     print r.get('name')

def logged(when):
    def log(f,*args,**kargs):
        print "called:function:%s args: %r kargs: %r"%(f,args,kargs)

    def pre_logged(f):
        def warpper(*args,**kargs):
            log(f,*args,**kargs)
            return f(*args,**kargs)
        return warpper
    def post_logged(f):
        def wrapper(*args,**kwargs):
            print "in the function (wrapper)"

            try:
                return f(*args,**kwargs)
            finally:
                log(f,*args,**kwargs)
                print "in the try-finally"
        return wrapper
    try:
        return {'pre':pre_logged,'post':post_logged}[when]
    except KeyError as e:
        print "must be pre or post"
@logged('post')
def hello(name):
    print 'hello,',name

if __name__=='__main__':
    hello('world')